Imports System.IO

Friend Class cServers

    Private Const sfName As String = "servers.xml"

    Friend oUp As ServerInfo
    Friend oDown As ServerInfo
    Friend oHeader As ServerInfo

    Public Function LoadServers() As Boolean

        oUp = Nothing
        oDown = Nothing
        oHeader = Nothing

        If Not System.IO.File.Exists(SettingsFolder() & "\" & sfName) Then

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
            doc.Load(SettingsFolder() & "\" & sfName)
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
            root = doc.CreateElement(Spotname)

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
                child.SetAttribute("SSL", sIIF(ZL.SSL, "1", "0"))
                child.SetAttribute("Connections", CStr(ZL.Connections))

                root.AppendChild(child)

            Next

            doc.AppendChild(root)

            If FileExists(SettingsFolder() & "\" & sfName) Then
                System.IO.File.SetAttributes(SettingsFolder() & "\" & sfName, IO.FileAttributes.Normal)
            End If

            doc.Save(SettingsFolder() & "\" & sfName)

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
