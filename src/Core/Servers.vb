Imports System.IO
Imports System.Xml

Imports Spotlib

Friend Class cServers

    Private Const sfName As String = "servers.xml"

    Friend oUp As ServerInfo
    Friend oDown As ServerInfo
    Friend oHeader As ServerInfo

    Public Function LoadServers() As Boolean

        oUp = Nothing
        oDown = Nothing
        oHeader = Nothing

        If Not System.IO.File.Exists(Tools.SettingsFolder() & "\" & sfName) Then

            oUp = GetOldUpload()
            oDown = GetOldDownload()
            oHeader = GetOldHeader()

            Return SaveServers()

        End If

        Dim KL As New cEncPass
        Dim root As Xml.XmlElement
        Dim doc As New Xml.XmlDocument
        Dim child2 As Xml.XmlElement
        Dim Zh As ServerInfo = Nothing

        Try

            doc.XmlResolver = Nothing
            doc.Load(Tools.SettingsFolder() & "\" & sfName)
            root = doc.DocumentElement

            For Each child2 In root

                Zh = New ServerInfo

                Zh.SSL = child2.GetAttribute("SSL").Trim = "1"

                Zh.Port = CInt(Val(child2.GetAttribute("Port")))
                Zh.Server = child2.GetAttribute("Server")

                If Microsoft.VisualBasic.Strings.Right(child2.GetAttribute("Password"), 1) <> "=" Then
                    Zh.Password = child2.GetAttribute("Password")
                Else
                    Zh.Password = KL.Decrypt(child2.GetAttribute("Password"))
                End If

                Zh.Username = child2.GetAttribute("Username")
                Zh.Connections = CInt(Val(child2.GetAttribute("Connections")))

                Select Case child2.GetAttribute("Type").Trim.ToUpper
                    Case "UPLOAD", "UPLOADS"
                        oUp = Zh

                    Case "DOWNLOAD", "DOWNLOADS"
                        oDown = Zh

                    Case "HEADER", "HEADERS"
                        oHeader = Zh

                End Select

            Next

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Public Function SaveServers(Optional ByRef sError As String = "") As Boolean

        Dim root As Xml.XmlElement
        Dim doc As New Xml.XmlDocument
        Dim KL As New cEncPass

        doc.XmlResolver = Nothing

        Try
            root = doc.CreateElement(Spotz.Spotname)

            Dim ZK As New List(Of ServerInfo)

            ZK.Add(oHeader)
            ZK.Add(oDown)
            ZK.Add(oUp)

            Dim Ci As Long = 0

            For Each ZL As ServerInfo In ZK
                Dim child As Xml.XmlElement = doc.CreateElement("Server")

                Ci += 1

                Select Case Ci
                    Case 1
                        child.SetAttribute("Type", "Headers")
                    Case 2
                        child.SetAttribute("Type", "Downloads")
                    Case 3
                        child.SetAttribute("Type", "Uploads")
                End Select

                child.SetAttribute("Server", ZL.Server)
                child.SetAttribute("Username", ZL.Username)

                If Len(ZL.Password) > 0 Then
                    child.SetAttribute("Password", KL.Encrypt(ZL.Password))
                Else
                    child.SetAttribute("Password", "")
                End If

                child.SetAttribute("Port", CStr(ZL.Port))
                child.SetAttribute("SSL", Utils.sIIF(ZL.SSL, "1", "0"))
                child.SetAttribute("Connections", CStr(ZL.Connections))

                root.AppendChild(child)

            Next

            doc.AppendChild(root)

            If Utils.FileExists(Tools.SettingsFolder() & "\" & sfName) Then
                System.IO.File.SetAttributes(Tools.SettingsFolder() & "\" & sfName, IO.FileAttributes.Normal)
            End If

            doc.Save(Tools.SettingsFolder() & "\" & sfName)

            Return True

        Catch ex As Exception

            sError = "SaveServers: " & ex.Message
            Return False

        End Try

    End Function

    Friend Function GetOldDownload() As ServerInfo

        Dim OD As New ServerInfo

        OD.Password = ""
        OD.Server = ""
        OD.Username = ""
        OD.Connections = 2
        OD.Port = 119
        OD.SSL = False

        Return OD

    End Function

    Friend Function GetOldHeader() As ServerInfo

        Dim OD As New ServerInfo

        OD.Password = ""
        OD.Server = ""
        OD.Username = ""
        OD.Connections = 2
        OD.Port = 119
        OD.SSL = False

        Return OD

    End Function

    Friend Function GetOldUpload() As ServerInfo

        Dim OD As New ServerInfo

        OD.Password = ""
        OD.Server = ""
        OD.Username = ""
        OD.Connections = 1
        OD.Port = 119
        OD.SSL = False

        Return OD

    End Function

End Class

Public Class ServerList

    Friend LastServer As String = ""
    Friend ServersDB As New cServers

    Private hPhuse As Phuse.Engine = Nothing
    Private dPhuse As Phuse.Engine = Nothing
    Private uPhuse As Phuse.Engine = Nothing

    Friend Function GetServer(ByVal Type As ServerType) As ServerInfo

        Select Case Type
            Case ServerType.Headers
                Return ServersDB.oHeader
            Case ServerType.Upload
                Return ServersDB.oUp
            Case ServerType.Download
                Return ServersDB.oDown
            Case Else
                Return Nothing
        End Select

    End Function

    Friend Sub CopyServer(ByVal sIn As ServerInfo, ByVal sOut As ServerInfo)

        With sIn
            sOut.SSL = .SSL
            sOut.Port = .Port
            sOut.Server = .Server
            sOut.Username = .Username
            sOut.Password = .Password
            sOut.Connections = .Connections
        End With

    End Sub

    Friend Sub ClearPhuses()

        If Not hPhuse Is Nothing Then

            Try
                hPhuse.Close()
            Catch
            End Try

            hPhuse = Nothing

        End If

        If Not dPhuse Is Nothing Then

            Try
                dPhuse.Close()
            Catch
            End Try

            dPhuse = Nothing

        End If

        If Not uPhuse Is Nothing Then

            Try
                uPhuse.Close()
            Catch
            End Try

            uPhuse = Nothing

        End If

    End Sub

    Friend Function HeaderPhuse() As Phuse.Engine

        If hPhuse Is Nothing Then

            If SameServer(GetServer(ServerType.Download), GetServer(ServerType.Headers)) Then
                hPhuse = DownloadPhuse()
            Else
                hPhuse = CreatePhuse(GetServer(ServerType.Headers))
            End If

        End If

        Return hPhuse

    End Function

    Friend Function UploadPhuse() As Phuse.Engine

        If uPhuse Is Nothing Then

            If SameServer(GetServer(ServerType.Download), GetServer(ServerType.Upload)) Then
                uPhuse = DownloadPhuse()
            Else
                uPhuse = CreatePhuse(GetServer(ServerType.Upload))
            End If

        End If

        Return uPhuse

    End Function

    Friend Function DownloadPhuse() As Phuse.Engine

        If dPhuse Is Nothing Then

            dPhuse = CreatePhuse(GetServer(ServerType.Download))

        End If

        Return dPhuse

    End Function

    Friend Function CreatePhuse(ByVal hServer As ServerInfo) As Phuse.Engine

        Dim hPhuse As New Phuse.Engine

        If Len(hServer.Server) = 0 Then Return Nothing
        hPhuse.Servers.Add(hServer.Server, hServer.Username, hServer.Password, hServer.Port, 1, hServer.SSL)

        Return hPhuse

    End Function

    Private Function SameServer(ByVal hServer As ServerInfo, ByVal cServer As ServerInfo) As Boolean

        If hServer.SSL <> cServer.SSL Then Return False
        If hServer.Port <> cServer.Port Then Return False
        If Trim(LCase(hServer.Server)) <> Trim(LCase(cServer.Server)) Then Return False
        If Trim(LCase(hServer.Username)) <> Trim(LCase(cServer.Username)) Then Return False

        Return True

    End Function

End Class


Friend Class XmlList

    Public Shared Function CreateKeys() As Boolean

        Dim sKey(10) As String

        Try

            If Not System.IO.File.Exists(Tools.SettingsFolder() & "\keys.xml") Then

                sKey(2) = "ys8WSlqonQMWT8ubG0tAA2Q07P36E+CJmb875wSR1XH7IFhEi0CCwlUzNqBFhC+P"
                sKey(3) = "uiyChPV23eguLAJNttC/o0nAsxXgdjtvUvidV2JL+hjNzc4Tc/PPo2JdYvsqUsat"
                sKey(4) = "1k6RNDVD6yBYWR6kHmwzmSud7JkNV4SMigBrs+jFgOK5Ldzwl17mKXJhl+su/GR9"

                Dim SK As New StreamWriter(Tools.SettingsFolder() & "\keys.xml", False, System.Text.Encoding.UTF8)

                SK.WriteLine("<Keys>")

                For Cl = 0 To 9
                    If Len(sKey(Cl)) > 0 Then
                        SK.WriteLine(vbTab & "<Key ID='" & Cl & "'>" & sKey(Cl) & "</Key>")
                    End If
                Next

                SK.WriteLine("</Keys>")

                SK.Close()

            End If

        Catch ex As Exception

            Tools.Foutje("Create_Keys: " & ex.Message)
            Return False

        End Try

        Return True

    End Function

    Public Shared Function CreateList(ByVal sFile As String, ByVal cList As List(Of ListItem)) As Boolean

        Try

            Dim SK As New StreamWriter(Tools.SettingsFolder() & "\" & sFile, False, System.Text.Encoding.UTF8)

            SK.WriteLine("<Keys>")

            For Each kItem As ListItem In cList
                If Len(kItem.Key) > 0 Then
                    If Len(kItem.Name) = 0 Then
                        SK.WriteLine(vbTab & "<Key>" & kItem.Key.Replace("<", "") & "</Key>")
                    Else
                        SK.WriteLine(vbTab & "<Key Name=" & Chr(34) & kItem.Name.Replace(Chr(34), "") & Chr(34) & ">" & kItem.Key.Replace("<", "") & "</Key>")
                    End If
                End If
            Next

            SK.WriteLine("</Keys>")
            SK.Close()

            Return True

        Catch ex As Exception

            Tools.Foutje("Create_List: " & ex.Message)

        End Try

        Return False

    End Function

    Friend Shared Function UpdateList(ByVal zUrl As String, ByVal sFile As String) As Boolean

        Try

            Dim sXml As String = Utils.DownloadString(zUrl, "")

            If sXml Is Nothing Then Return False
            If Len(sXml) = 0 Then Return False
            If Not ValidList(sXml) Then Return False

            Dim SK As New StreamWriter(Tools.SettingsFolder() & "\" & sFile, False, System.Text.Encoding.UTF8)
            SK.Write(sXml)
            SK.Close()

            Return True

        Catch
        End Try

        Return False

    End Function

    Friend Shared Function ValidList(ByVal sXML As String) As Boolean

        Try

            Dim root As XmlElement
            Dim doc As New XmlDocument

            doc.XmlResolver = Nothing
            doc.LoadXml(sXML)
            root = doc.DocumentElement

            If root.Name.ToLower <> "keys" Then
                Throw New Exception("XML Error")
            End If

            Return True

        Catch
        End Try

        Return False

    End Function

    Friend Shared Function LoadList(ByVal sFile As String, ByVal ToList As HashSet(Of String)) As Boolean

        Try

            Dim root As XmlElement
            Dim doc As New XmlDocument

            doc.XmlResolver = Nothing
            doc.Load(Tools.SettingsFolder() & "\" & sFile)
            root = doc.DocumentElement

            If root.Name.ToLower <> "keys" Then
                Throw New Exception("XML Error")
            End If

            For Each child2 As XmlElement In root
                ToList.Add(child2.InnerText)
            Next

            Return True

        Catch ex As Exception
            Throw New Exception("Load_List: " & ex.Message)

        End Try

        Return False

    End Function

    Friend Shared Sub UpdateBlacklist(iBlackList As HashSet(Of String))

        Dim sServer As String = "blacklist.server.xml"

        If Len(My.Settings.BlacklistURL) > 0 Then

            UpdateList(Utils.AddHttp(My.Settings.BlacklistURL), sServer)

            If System.IO.File.Exists(Tools.SettingsFolder() & "\" & sServer) Then

                LoadList(sServer, iBlackList)

            End If

        End If

    End Sub

    Public Shared Function AddToList(ByVal sName As String, ByVal sKey As String, ByVal sFile As String) As Boolean

        Try

            Dim root As XmlElement
            Dim doc As New XmlDocument

            doc.XmlResolver = Nothing
            doc.Load(Tools.SettingsFolder() & "\" & sFile)
            root = doc.DocumentElement

            If root.Name.ToLower <> "keys" Then
                Throw New Exception("XML Error")
            End If

            For Each child2 As XmlElement In root
                If child2.InnerText.ToLower.Trim = sKey.ToLower.Trim Then
                    Return True
                End If
            Next

            Dim NC As XmlNode = doc.CreateElement("Key")
            Dim NA As XmlAttribute = doc.CreateAttribute("Name")

            NA.Value = sName
            NC.Attributes.Append(NA)
            NC.InnerText = sKey

            root.AppendChild(NC)

            If Utils.FileExists(Tools.SettingsFolder() & "\" & sFile) Then
                System.IO.File.SetAttributes(Tools.SettingsFolder() & "\" & sFile, IO.FileAttributes.Normal)
            End If

            doc.Save(Tools.SettingsFolder() & "\" & sFile)

            Return True

        Catch
        End Try

        Return False

    End Function

    Public Shared Function RemoveFromList(ByVal sKey As String, ByVal sFile As String) As Boolean

        Try

            Dim root As XmlElement
            Dim doc As New XmlDocument
            Dim bDidSome As Boolean = False

            If Not Utils.FileExists(Tools.SettingsFolder() & "\" & sFile) Then Return False

            doc.XmlResolver = Nothing
            doc.Load(Tools.SettingsFolder() & "\" & sFile)
            root = doc.DocumentElement

            If root.Name.ToLower <> "keys" Then
                Throw New Exception("XML Error")
            End If

            For Each child2 As XmlElement In root
                If child2.InnerText.ToLower.Trim = sKey.ToLower.Trim Then
                    bDidSome = True
                    root.RemoveChild(child2)
                End If
            Next

            If Not bDidSome Then Return False

            If Utils.FileExists(Tools.SettingsFolder() & "\" & sFile) Then
                System.IO.File.SetAttributes(Tools.SettingsFolder() & "\" & sFile, IO.FileAttributes.Normal)
            End If

            doc.Save(Tools.SettingsFolder() & "\" & sFile)

            Return True

        Catch
        End Try

        Return False

    End Function

    Friend Shared Sub UpdateWhitelist(iWhitelist As HashSet(Of String))

        Dim sServer As String = "whitelist.server.xml"

        If Len(My.Settings.WhitelistURL) > 0 Then

            UpdateList(Utils.AddHttp(My.Settings.WhitelistURL), sServer)

            If System.IO.File.Exists(Tools.SettingsFolder() & "\" & sServer) Then

                LoadList(sServer, iWhitelist)

            End If

        End If

    End Sub

    Friend Shared Function LoadKeys() As String()

        Static sKey(10) As String
        Static KeysLoaded As Boolean = False

        If KeysLoaded Then Return sKey

        Try

            Dim sLocal As String = "keys.xml"

            If Len(My.Settings.KeysURL) > 0 Then

                UpdateList(Utils.AddHttp(My.Settings.KeysURL), sLocal)

            End If

            If Not System.IO.File.Exists(Tools.SettingsFolder() & "\" & sLocal) Then

                CreateKeys()

            End If

            Dim root As XmlElement
            Dim doc As New XmlDocument
            Dim child2 As XmlElement

            doc.XmlResolver = Nothing
            doc.Load(Tools.SettingsFolder() & "\" & sLocal)
            root = doc.DocumentElement

            For Each child2 In root
                If Not child2.GetAttribute("ID") Is Nothing Then
                    If CInt(Val(child2.GetAttribute("ID"))) >= 2 Then
                        If CInt(Val(child2.GetAttribute("ID"))) <= 8 Then
                            If Not child2.InnerXml.ToLower.Contains("rsakeyvalue") Then
                                sKey(CInt(Val(child2.GetAttribute("ID")))) = child2.InnerText
                            Else
                                sKey(CInt(Val(child2.GetAttribute("ID")))) = child2.ChildNodes(0).ChildNodes(0).InnerText
                            End If
                        End If
                    End If
                End If
            Next

            KeysLoaded = True
            Return sKey

        Catch ex As Exception

            Throw New Exception("Load_Keys: " & ex.Message)

        End Try

    End Function

End Class
