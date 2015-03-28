Imports System.Xml
Imports System.IO
Imports System.Threading
Imports System.Windows.Threading
Imports System.Text
Imports System.Globalization
Imports System.Net
Imports System.Windows.Media.Imaging
Imports Spotnet.Spotnet
Imports System.ComponentModel

Friend Class SabHelper

    Public SabStarted As Boolean = False
    Private SabUpdate As DispatcherTimer

    Public WithEvents SabCol As New SabItems
    Private WithEvents SabUpdater As BackgroundWorker

    Public Event ProgressChanged(ByVal lVal As Integer)

    Public Function GetSabIni(ByVal Srv As ServerInfo) As String

        Dim objReader As StreamWriter

        If Not FileExists(SettingsFolder() & "\sabnzbd.ini") Then
            Try
                objReader = New StreamWriter(SettingsFolder() & "\sabnzbd.ini", False, LatinEnc)
                objReader.Write(My.Resources.sabnzbd.ToString())
                objReader.Close()
            Catch Ex As Exception
                Return vbNullString
            End Try
        End If

        Dim zIni As String = GetFileContents(SettingsFolder() & "\sabnzbd.ini")

        If Len(zIni) = 0 Then Return vbNullString
        zIni = zIni.Replace("[HOST]", CStr(My.Settings.SabHost)).Replace("[PORT]", CStr(My.Settings.SabPort))

        Dim xIni() As String = zIni.Replace(vbCr, "").Split(vbLf.ToCharArray())

        For XH As Integer = 0 To UBound(xIni)

            If xIni(XH).Contains("[ROOT]") Then
                xIni(XH) = xIni(XH).Replace("[ROOT]", "'''" & DownDir())
                xIni(XH) += "'''"
            End If

            If xIni(XH).Contains("[ROOT2]") Then
                xIni(XH) = xIni(XH).Replace("[ROOT2]", "'''" & SettingsFolder())
                xIni(XH) += "'''"
            End If

        Next

        zIni = Join(xIni, vbCrLf)

        If Len(xIni(UBound(xIni))) > 0 Then
            zIni += vbCrLf
        End If

        If Not zIni.ToLower.Contains("[servers]") Then

            LastServer = Srv.Server.ToLower & ":" & Srv.Port

            zIni += "[servers]" & vbCrLf & _
            "[[" & LastServer & "]]" & vbCrLf & _
            "username = '''" & Srv.Username & "'''" & vbCrLf & _
            "enable = 1" & vbCrLf & _
            "name = " & LastServer & vbCrLf & _
            "fillserver = 0" & vbCrLf & _
            "connections = " & Srv.Connections & vbCrLf & _
            "ssl = " & sIIF(Srv.SSL, "1", "0") & vbCrLf & _
            "host = " & Srv.Server.ToLower & vbCrLf & _
            "timeout = 120" & vbCrLf & _
            "password = '''" & Srv.Password & "'''" & vbCrLf & _
            "optional = 0" & vbCrLf & _
            "port = " & Srv.Port & vbCrLf

        End If

        Dim sFile As String = System.IO.Path.GetTempFileName & ".ini"

        Try
            objReader = New StreamWriter(sFile, False, LatinEnc)
            objReader.Write(zIni)
            objReader.Close()
        Catch Ex As Exception
            Return vbNullString
        End Try

        Return sFile

    End Function

    Private Function PostData(ByVal sCmd As String, ByVal zPostData As String, ByRef zError As String) As String

        Dim sLoc As String = ""
        Dim sReturn As String = ""
        Dim bytArguments As Byte()
        Dim oWeb As System.Net.WebClient
        Dim zR As New Random

        Try

            oWeb = New System.Net.WebClient
            oWeb.Headers.Add("Content-Type", "application/x-www-form-urlencoded")

            bytArguments = MakeLatin(zPostData)

            Try

                sReturn = GetLatin(oWeb.UploadData(GetSabURL() & sCmd & "?_dc=" & CStr(zR.Next(0, 999999)) & "&session=xxx", "POST", bytArguments))

            Catch ex As Exception

                zError = "Kan de server niet bereiken. (" & ex.Message & ")"
                Return vbNullString

            End Try

            oWeb.Dispose()
            oWeb = Nothing

            zError = zPostData
            Return sReturn

        Catch ex As Exception
            zError = ex.Message
            Return vbNullString
        End Try

    End Function

    Public Function UpdateFolder(ByRef sError As String) As Boolean

        If External() Then Return True

        Dim sReq As String
        sReq = "download_dir=" & URLEncode(DownDir() & "\Incompleet") & "&download_free=&complete_dir=" & URLEncode(DownDir()) & "&dirscan_dir=&dirscan_speed=5&script_dir=&email_dir=&cache_dir=" & URLEncode(SettingsFolder() & "\cache") & "&log_dir=" & URLEncode(SettingsFolder() & "\logs") & "&nzb_backup_dir="

        Return PostData("config/directories/saveDirectories", sReq, sError).ToLower.Contains("javascript:submitconfig(")

    End Function

    Public Function UpdateSetting(ByVal Srv As ServerInfo, ByRef sError As String) As Boolean

        If External() Then Return True
        If LastServer.Length = 0 Then Return True

        Dim sReq As String
        Dim sHost As String = Srv.Server

        sReq = "host=" & URLEncode(sHost.ToLower) & "&port=" & CStr(Srv.Port) & "&username=" & URLEncode(Srv.Username) & "&password=" & URLEncode(Srv.Password) & "&timeout=120&connections=" & Srv.Connections & "&server=" & URLEncode(LastServer) & "&ssl=" & sIIF(Srv.SSL, "1", "0") & "&enable=1" '' 

        If Not PostData("config/server/saveServer", sReq, sError).ToLower.Contains("javascript:testserver(") Then Return False

        LastServer = sHost.ToLower & ":" & Srv.Port
        Return True

    End Function

    Public Function IsSabRunning(ByVal mTime As Integer) As Boolean

        Try

            If SabStarted Then Return True

            Dim VR As HttpWebRequest = CType(WebRequest.Create(GetSabURL() & "sabnzbd/api?" & "mode=version&output=xml" & GetAuth()), HttpWebRequest)

            VR.Proxy = Nothing
            VR.Timeout = mTime
            VR.KeepAlive = False

            Dim VrResp As HttpWebResponse = CType(VR.GetResponse, HttpWebResponse)
            Dim s As Stream = VrResp.GetResponseStream
            Dim objStreamRead As StreamReader = New System.IO.StreamReader(s, System.Text.Encoding.GetEncoding("utf-8"))
            Dim zXML As String = objStreamRead.ReadToEnd()

            If Len(zXML) = 0 Then Return False

            Dim m_xmld As XmlDocument
            m_xmld = New XmlDocument()
            m_xmld.XmlResolver = Nothing

            m_xmld.LoadXml(zXML)

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function AddNZB(ByVal sFile As String, ByVal sTitle As String, ByRef zError As String) As Boolean

        Dim xUrl As String

        Try

            If Not SabStarted Then Return False

            sTitle = Trim(MakeFilename(sTitle))

            If Len(sTitle) = 0 Then sTitle = "Leeg"
            If Len(sTitle) > 140 Then sTitle = Microsoft.VisualBasic.Left(sTitle, 140)

            Dim zR As New Random
            Dim sName As String = sTitle & "." & CStr(zR.Next(0, 999999)) & ".nzb"

            xUrl = GetSabURL() & "sabnzbd/api?mode=addfile&name=" & URLEncode(sName) & "&nzbname=" & URLEncode(sTitle) & GetAuth()

            Dim WebRequest As Net.HttpWebRequest = CType(Net.HttpWebRequest.Create(xUrl), Net.HttpWebRequest)

            WebRequest.Proxy = Nothing
            WebRequest.Method = "POST"
            WebRequest.KeepAlive = True

            Dim Boundary As String = "----------------------------" + DateTime.Now.Ticks.ToString("x")
            WebRequest.ContentType = "multipart/form-data; boundary=" + Boundary

            Dim sb As New StringBuilder

            sb.Append(vbCrLf & "--" & Boundary & vbCrLf & "Content-Disposition: form-data; name=""name""; filename=" & Chr(34) & sName & Chr(34))
            sb.Append(vbCrLf)
            sb.Append("Content-Type: application/x-nzb")
            sb.Append(vbCrLf)
            sb.Append(vbCrLf)

            Dim postHeaderBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(sb.ToString)
            Dim boundaryBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(vbCrLf & "--" & Boundary & "--" & vbCrLf)

            Dim filestream As New FileStream(sFile, IO.FileMode.Open, IO.FileAccess.Read)
            Dim length As Long = postHeaderBytes.Length + filestream.Length + boundaryBytes.Length
            WebRequest.ContentLength = length

            Dim requestStream As System.IO.Stream = WebRequest.GetRequestStream
            requestStream.Write(postHeaderBytes, 0, postHeaderBytes.Length)

            Dim buffer(8192) As Byte
            Dim bytesRead As Integer = -1

            While bytesRead <> 0
                bytesRead = filestream.Read(buffer, 0, buffer.Length)
                requestStream.Write(buffer, 0, bytesRead)
            End While

            requestStream.Write(boundaryBytes, 0, boundaryBytes.Length)

            Dim Response As Net.WebResponse = WebRequest.GetResponse
            Dim s As Stream = Response.GetResponseStream
            Dim objStreamRead As StreamReader = New System.IO.StreamReader(s, System.Text.Encoding.GetEncoding("utf-8"))

            Dim zXML As String = objStreamRead.ReadToEnd()

            If Len(zXML) = 0 Then Return False
            If zXML.Trim.ToLower <> "ok" Then zError = zXML : Return False

            objStreamRead.Close()
            s.Close()
            Response.Close()
            requestStream.Close()
            filestream.Close()

            Return True

        Catch ex As Exception
            zError = ex.Message
            Return False
        End Try

    End Function

    Public Function DoQueueCmd(ByVal sID As String, ByVal sCmd As String, ByVal sHistory As Boolean, ByRef zError As String) As Boolean

        Dim zXML As String

        Try

            If Not SabStarted Then Return False

            zXML = GetSabCMD("mode=" & sIIF(Not sHistory, "queue", "history") & "&name=" & sCmd & "&value=" & sID, zError)

            If Len(zXML) = 0 Then Return False
            If zXML.Trim.ToLower <> "ok" Then zError = zXML : Return False

            Return True

        Catch ex As Exception
            zError = ex.Message
            Return False
        End Try

    End Function

    Public Function DoSwitchCmd(ByVal sVal1 As String, ByVal sVal2 As String, ByRef zError As String) As Boolean

        Dim zXML As String

        Try

            If Not SabStarted Then Return False

            zXML = GetSabCMD("mode=switch&value=" & sVal1 & "&value2=" & sVal2, zError)

            If Len(zXML) = 0 Then Return False

            Return True

        Catch ex As Exception
            zError = ex.Message
            Return False
        End Try

    End Function

    Public Function GetSabQueue(ByRef xParts As List(Of String()), ByRef zError As String) As Boolean

        Dim zXML As String
        Dim KBPerSec As String
        Dim xDataSlots As New List(Of String())

        Try

            zXML = GetSabCMD("mode=queue&output=xml", zError)
            If Len(zXML) = 0 Then Return False

            Dim m_xmld As XmlDocument

            m_xmld = New XmlDocument()
            m_xmld.XmlResolver = Nothing

            m_xmld.LoadXml(zXML)

            KBPerSec = m_xmld.DocumentElement.SelectSingleNode("kbpersec").InnerText

            For Each XZ As XmlNode In m_xmld.DocumentElement.SelectSingleNode("slots").SelectNodes("slot")

                Dim SlotData(8) As String

                With XZ

                    SlotData(0) = .SelectSingleNode("filename").InnerText
                    SlotData(1) = .SelectSingleNode("status").InnerText
                    ''SlotData(3) = .SelectSingleNode("eta").InnerText
                    SlotData(2) = .SelectSingleNode("timeleft").InnerText
                    ''SlotData(5) = .SelectSingleNode("avg_age").InnerText
                    ''SlotData(6) = .SelectSingleNode("script").InnerText
                    ''SlotData(7) = .SelectSingleNode("msgid").InnerText
                    ''SlotData(8) = .SelectSingleNode("verbosity").InnerText
                    ''SlotData(9) = .SelectSingleNode("mb").InnerText
                    ''SlotData(10) = .SelectSingleNode("mbleft").InnerText
                    ''SlotData(11) = .SelectSingleNode("filename").InnerText
                    ''SlotData(12) = .SelectSingleNode("priority").InnerText
                    ''SlotData(13) = .SelectSingleNode("cat").InnerText
                    SlotData(3) = .SelectSingleNode("percentage").InnerText
                    SlotData(4) = .SelectSingleNode("nzo_id").InnerText
                    ''SlotData(16) = .SelectSingleNode("unpackopts").InnerText
                    SlotData(5) = .SelectSingleNode("size").InnerText
                    SlotData(6) = .SelectSingleNode("index").InnerText
                    SlotData(7) = ""
                    SlotData(8) = KBPerSec

                    xDataSlots.Add(SlotData)
                End With

            Next

            xParts = xDataSlots

            Return True

        Catch ex As Exception
            zError = ex.Message
            Return False
        End Try

    End Function

    Private Function HighestIndex() As Long

        Dim xHigh As Long = 0

        For Each X As SabItem In SabCol
            If Not X.History Then
                If X.Index > xHigh Then xHigh = X.Index
            End If
        Next

        Return xHigh

    End Function

    Private Function NextItem(ByVal zzIndex As Long) As String

        For i As Integer = 0 To SabCol.Count - 2
            Dim XS As SabItem = SabCol.Item(i)
            If XS.Index = zzIndex Then
                Dim XS2 As SabItem = SabCol.Item(i + 1)
                Return XS2.ID
            End If
        Next

        Return vbNullString

    End Function

    Private Function PrevItem(ByVal zzIndex As Long) As String

        For i As Integer = 1 To SabCol.Count - 1
            Dim XS As SabItem = SabCol.Item(i)
            If XS.Index = zzIndex Then
                Dim XS2 As SabItem = SabCol.Item(i - 1)
                Return XS2.ID
            End If
        Next

        Return vbNullString

    End Function

    Private Sub DoOpen(ByVal sender As Object, ByVal e As RoutedEventArgs)

        Dim zErr As String = ""

        Dim x As MenuItem = CType(e.OriginalSource, MenuItem)
        Dim x2 As ContextMenu = CType(x.Parent, ContextMenu)
        Dim x3 As SabItem = CType(x2.Tag, SabItem)
        Dim OldIndex As Long

        OldIndex = x3.Index

        Select Case CStr(x.Tag)
            Case "OPEN"

                DoOpen(x3)

            Case "UP"
                If DoSwitchCmd(x3.ID, PrevItem(OldIndex), zErr) Then
                    SabCol.RemoveID(x3.ID)
                Else
                    Foutje(zErr)
                End If
            Case "DOWN"
                If DoSwitchCmd(x3.ID, NextItem(OldIndex), zErr) Then
                    SabCol.RemoveID(x3.ID)
                Else
                    Foutje(zErr)
                End If
            Case "PAUSE"
                If Not DoQueueCmd(x3.ID, "pause", x3.IsHistory, zErr) Then
                    Foutje(zErr)
                End If
            Case "RESUME"
                If Not DoQueueCmd(x3.ID, "resume", x3.IsHistory, zErr) Then
                    Foutje(zErr)
                End If
            Case "DELETE"

                DeleteDownload(x3)

        End Select

        UpdateSabAsync()

    End Sub

    Private Sub DoOpenMulti(ByVal sender As Object, ByVal e As RoutedEventArgs)

        Dim zErr As String = ""

        Dim x As MenuItem = CType(e.OriginalSource, MenuItem)
        Dim x2 As ContextMenu = CType(x.Parent, ContextMenu)
        Dim x3 As List(Of SabItem) = CType(x2.Tag, List(Of SabItem))

        For Each x4 As SabItem In x3

            Select Case CStr(x.Tag)

                Case "PAUSE"

                    If Not x4.IsHistory Then
                        If Not x4.IsPaused Then
                            If Not DoQueueCmd(x4.ID, "pause", x4.IsHistory, zErr) Then
                                Foutje(zErr)
                            End If
                        End If
                    End If

                Case "RESUME"

                    If Not x4.IsHistory Then
                        If x4.IsPaused Then
                            If Not DoQueueCmd(x4.ID, "resume", x4.IsHistory, zErr) Then
                                Foutje(zErr)
                            End If
                        End If
                    End If

                Case "DELETE"

                    DeleteDownload(x4)

            End Select

        Next

        UpdateSabAsync()

    End Sub

    Public Sub DeleteDownload(ByVal x4 As SabItem)

        Dim zErr As String = ""

        If Not DoQueueCmd(x4.ID, "delete", x4.IsHistory, zErr) Then
            Foutje(zErr)
        End If

        If My.Computer.Keyboard.ShiftKeyDown Or My.Settings.RemoveDownloads Then

            If x4.IsHistory Then
                If x4.Location.Length > 0 Then
                    If Not DirectoryExists(x4.Location) Then
                        If FileExists(x4.Location) Then
                            x4.Location = Path.GetDirectoryName(x4.Location)
                        End If
                    End If
                    If DirectoryExists(x4.Location) Then
                        My.Computer.FileSystem.DeleteDirectory(x4.Location, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    End If
                Else
                    ' TODO: Not supported by Sab
                End If
            Else
                ' TODO: Not supported by Sab
            End If

        End If

    End Sub

    Public Function DoOpen(ByRef zSend As SabItem) As Boolean

        If CanOpen(zSend) Then
            Process.Start("explorer.exe", zSend.Location)
            Return True
        End If

        Return False

    End Function

    Private Function CanOpen(ByRef zSend As SabItem) As Boolean

        If zSend.IsHistory Then
            If zSend.Location.Length > 0 Then
                If Not DirectoryExists(zSend.Location) Then
                    If FileExists(zSend.Location) Then
                        zSend.Location = Path.GetDirectoryName(zSend.Location)
                    End If
                End If
                If DirectoryExists(zSend.Location) Then
                    Return True
                End If
            End If
        End If

        Return False

    End Function

    Public Function GetMenu(ByRef zSend As SabItem) As ContextMenu

        Dim XM As New ContextMenu
        Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)

        XM.FontFamily = Ref.FontFamily
        XM.FontSize = Ref.FontSize - 1
        XM.FontStyle = Ref.FontStyle

        XM.Tag = zSend

        Dim mOpen As New MenuItem
        With mOpen
            .Header = "Openen"
            .Tag = "OPEN"
            .IsEnabled = CanOpen(zSend)
            .Icon = GetIcon("open")
            If Not .IsEnabled Then .Opacity = 0.5
            .AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf DoOpen))
        End With

        Dim mOmhoog As New MenuItem
        With mOmhoog
            .Header = "Omhoog"
            .Tag = "UP"
            .IsEnabled = (zSend.Index > 0) And (Not zSend.IsHistory)
            .Icon = GetIcon("up")
            If Not .IsEnabled Then .Opacity = 0.5
            .AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf DoOpen))
        End With

        Dim mOmlaag As New MenuItem
        With mOmlaag
            .Header = "Omlaag"
            .Tag = "DOWN"
            .IsEnabled = (zSend.Index < HighestIndex()) And (Not zSend.IsHistory)
            .Icon = GetIcon("down")
            If Not .IsEnabled Then .Opacity = 0.5
            .AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf DoOpen))
        End With

        Dim mPause As New MenuItem
        With mPause
            If Not zSend.IsPaused Then
                .Header = "Pauzeren"
                .Tag = "PAUSE"
                .Icon = GetIcon("pause")
            Else
                .Header = "Hervatten"
                .Tag = "RESUME"
                .Icon = GetIcon("resume")
            End If
            .IsEnabled = (Not zSend.IsHistory)
            If Not .IsEnabled Then .Opacity = 0.5
            .AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf DoOpen))
        End With

        Dim mDelete As New MenuItem
        With mDelete
            .Header = "Verwijderen"
            .Tag = "DELETE"
            .Icon = GetIcon("delete")
            If Not .IsEnabled Then .Opacity = 0.5
            .AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf DoOpen))
        End With

        XM.Items.Add(mOpen)

        If (mOmhoog.IsEnabled Or mOmlaag.IsEnabled) Then
            XM.Items.Add(New Separator)
            XM.Items.Add(mOmhoog)
            XM.Items.Add(mOmlaag)
        End If

        XM.Items.Add(New Separator)
        XM.Items.Add(mPause)
        XM.Items.Add(mDelete)

        Return XM

    End Function

    Public Function GetMultiMenu(ByRef zSend As List(Of SabItem)) As ContextMenu

        Dim XM As New ContextMenu
        Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)

        XM.FontFamily = Ref.FontFamily
        XM.FontSize = Ref.FontSize - 1
        XM.FontStyle = Ref.FontStyle

        XM.Tag = zSend

        Dim ShowPause As Boolean = False
        Dim ShowResume As Boolean = False

        For Each x4 As SabItem In zSend
            If Not x4.IsHistory Then
                If x4.IsPaused Then
                    ShowResume = True
                Else
                    ShowPause = True
                End If
            End If
        Next

        Dim mDelete As New MenuItem
        With mDelete
            .Header = "Verwijderen"
            .Tag = "DELETE"
            .Icon = GetIcon("delete")
            If Not .IsEnabled Then .Opacity = 0.5
            .AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf DoOpenMulti))
        End With

        Dim mPause As New MenuItem
        With mPause
            .Header = "Pauzeren"
            .Tag = "PAUSE"
            .IsEnabled = ShowPause
            .Icon = GetIcon("pause")
            If Not .IsEnabled Then .Opacity = 0.5
            .AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf DoOpenMulti))
        End With

        Dim mResume As New MenuItem
        With mResume
            .Header = "Hervatten"
            .Tag = "RESUME"
            .IsEnabled = ShowResume
            .Icon = GetIcon("resume")
            If Not .IsEnabled Then .Opacity = 0.5
            .AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf DoOpenMulti))
        End With

        If ShowPause Or ShowResume Then

            If ShowPause Then XM.Items.Add(mPause)
            If ShowResume Then XM.Items.Add(mResume)

            XM.Items.Add(New Separator)

        End If

        XM.Items.Add(mDelete)

        Return XM

    End Function

    Public Function GetSabHistory(ByRef xParts As List(Of String()), ByRef zError As String) As Boolean

        Dim zXML As String
        Dim xDataSlots As New List(Of String())

        Try

            zXML = GetSabCMD("mode=history&output=xml", zError)
            If Len(zXML) = 0 Then Return False

            Dim m_xmld As XmlDocument

            m_xmld = New XmlDocument()
            m_xmld.XmlResolver = Nothing

            m_xmld.LoadXml(zXML)

            For Each XZ As XmlNode In m_xmld.DocumentElement.SelectSingleNode("slots").SelectNodes("slot")

                Dim SlotData(8) As String

                With XZ

                    SlotData(0) = .SelectSingleNode("name").InnerText
                    SlotData(1) = .SelectSingleNode("status").InnerText
                    SlotData(2) = ""
                    SlotData(3) = "100"
                    SlotData(4) = .SelectSingleNode("nzo_id").InnerText
                    SlotData(5) = .SelectSingleNode("size").InnerText
                    SlotData(6) = CStr(99)
                    SlotData(7) = .SelectSingleNode("storage").InnerText
                    SlotData(8) = "0.00"

                    xDataSlots.Add(SlotData)
                End With

            Next

            xParts = xDataSlots

            Return True

        Catch ex As Exception
            zError = ex.Message
            Return False
        End Try

    End Function

    Private Function GetSabURL() As String

        Return "http://" & My.Settings.SabHost & ":" & CStr(My.Settings.SabPort) & "/"

    End Function

    Public Function GetSabCMD(ByVal zUrl As String, ByRef zError As String) As String

        Dim oWeb As System.Net.WebClient

        Try

            oWeb = New System.Net.WebClient
            Return (oWeb.DownloadString(GetSabURL() & "sabnzbd/api?" & zUrl & GetAuth()))

        Catch ex As Exception
            zError = ex.Message
            Return vbNullString
        End Try

    End Function

    Private Sub ResumeSab()

        Dim CS As New SabHelper
        Dim oWeb As System.Net.WebClient

        Try

            oWeb = New System.Net.WebClient
            oWeb.DownloadStringAsync(New System.Uri(CS.GetSabURL & "sabnzbd/api?mode=resume" & CS.GetAuth()))

        Catch ex As Exception
        End Try

    End Sub

    Public Function IsProcessRunning() As Boolean

        Try

            Dim SabPath As String = AppPath()

            For Each clsProcess As Process In Process.GetProcesses()
                If Not clsProcess Is Nothing Then
                    If Not clsProcess.ProcessName Is Nothing Then
                        If clsProcess.ProcessName.ToLower = "sabnzbd" Then
                            If Path.GetDirectoryName(clsProcess.MainModule.FileName).ToLower = Path.GetDirectoryName(SabPath & "\sabnzbd.exe").ToLower Then
                                Return True
                            End If
                        End If
                    End If
                End If
            Next

            Return False

        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function CloseSab(Optional ByRef zError As String = "") As Boolean

        Try

            If Not SabStarted Then
                If Not IsProcessRunning() Then Return True
            End If

            If Not SabUpdate Is Nothing Then
                SabUpdate.Stop()
                SabUpdate.IsEnabled = False
                SabUpdate = Nothing
            End If

            SabCol.ClearAll() ' To clear the grid
            SabCol.Clear()

            If Not External() Then

                Dim CS As New SabHelper
                Dim oWeb As System.Net.WebClient

                Try

                    oWeb = New System.Net.WebClient
                    oWeb.DownloadString(New System.Uri(CS.GetSabURL & "sabnzbd/api?mode=shutdown" & CS.GetAuth()))

                Catch ex As Exception
                End Try

            End If

            SabStarted = False

            Return True

        Catch ex As Exception
            zError = Err.Description
            Return False
        End Try

    End Function

    Public Function LoadSab(ByVal Srv As ServerInfo, Optional ByRef zError As String = "") As Boolean

        If SabStarted Then Return True
        If IsProcessRunning() Then Return True

        Try

            Dim IniFile As String = GetSabIni(Srv)

            If Len(IniFile) = 0 Then
                zError = "Kan SAB ini file niet aanmaken?"
                Return False
            End If

            Dim pStart As New System.Diagnostics.Process
            Dim startInfo As System.Diagnostics.ProcessStartInfo
            Dim SabPath As String = AppPath()

            startInfo = New System.Diagnostics.ProcessStartInfo(SabPath & "\SABnzbd.exe")

            startInfo.Arguments = "-d -f " & Chr(34) & IniFile & Chr(34)
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden

            startInfo.CreateNoWindow = True
            startInfo.UseShellExecute = False

            pStart.StartInfo = startInfo

            If Not pStart.Start() Then
                zError = "Kan SABnzbd niet starten?"
                Return False
            End If

            pStart.Dispose()
            pStart = Nothing

            Return True

        Catch ex As Exception
            zError = ex.Message
        End Try

        Return False

    End Function

    Public Function StartSab(ByVal Srv As ServerInfo, ByVal SkipPolling As Boolean, Optional ByRef zError As String = "") As Boolean

        If SabStarted Then Return SabStarted

        Try

            If Not External() Then

                If Not LoadSab(Srv, zError) Then Return False

            End If

            If Not SkipPolling Then

                Dim xStart As Date = Now
                Dim xDelay As Integer = 100

                Do While Not IsSabRunning(xDelay)
                    Wait(xDelay)
                    xDelay = xDelay + 30
                    System.Windows.Forms.Application.DoEvents()
                    If DateDiff("s", xStart, Now) > 11 Then
                        zError = "Bekijk de logfile."
                        Return False
                    End If
                Loop

            End If

            ResumeSab()

            SabUpdate = New DispatcherTimer(DispatcherPriority.Background)
            AddHandler SabUpdate.Tick, AddressOf DoUpdate

            SabUpdate.Interval = TimeSpan.FromMilliseconds(10)
            SabUpdate.Start()

            SabStarted = True
            Return SabStarted

        Catch ex As Exception
            zError = ex.Message
        End Try

        Return False

    End Function

    Private Sub SetInterval(ByVal iFactor As Integer)

        If Not SabUpdate Is Nothing Then
            If Not External() Then
                SabUpdate.Interval = TimeSpan.FromMilliseconds(My.Settings.SabRefresh * iFactor)
            Else
                SabUpdate.Interval = TimeSpan.FromMilliseconds((My.Settings.SabRefresh * 2) * iFactor)
            End If
        End If

    End Sub

    Public Sub DoUpdate(ByVal sender As Object, ByVal e As System.EventArgs)

        Static DidOnce As Boolean = False

        Try

            Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)

            If Ref.DownloadsVisible Then
                SetInterval(1)
            Else
                SetInterval(5)
            End If

            UpdateSabAsync()

        Catch ex As Exception

            If Not DidOnce Then
                DidOnce = True
                Foutje("Sab.DoUpdate: " & ex.Message)
            End If

        End Try

    End Sub

    Public Sub QuickRefresh()

        If Not SabUpdate Is Nothing Then SabUpdate.Interval = TimeSpan.FromMilliseconds(10)

    End Sub

    Public Function HasQueue() As Boolean

        Dim xDataSlots As New List(Of String())

        If Not SabStarted Then Return False

        If GetSabQueue(xDataSlots, "") Then

            Return xDataSlots.Count > 0

        Else

            Return True

        End If

    End Function

    Public Sub UpdateSabAsync()

        If SabUpdater Is Nothing Then

            SabUpdater = New BackgroundWorker()
            SabUpdater.RunWorkerAsync()

        End If

    End Sub

    Private Function GetApiKey() As String

        If Len(My.Settings.SabAPI) > 0 Then
            If My.Settings.SabAPI.Trim.Length > 0 Then
                Return "&apikey=" & My.Settings.SabAPI.Trim
            End If
        End If

        Return ""

    End Function

    Private Function GetAuth() As String

        Return GetSabUser() & GetSabPass() & GetApiKey()

    End Function

    Private Function GetSabPass() As String

        'If Len(My.Settings.SabPass) > 0 Then
        '    If My.Settings.SabPass.Trim.Length > 0 Then
        '        Return "&ma_password=" & My.Settings.SabPass.Trim
        '    End If
        'End If

        Return ""

    End Function

    Private Function GetSabUser() As String

        'If Len(My.Settings.SabUser) > 0 Then
        '    If My.Settings.SabUser.Trim.Length > 0 Then
        '        Return "&ma_username=" & My.Settings.SabUser.Trim
        '    End If
        'End If

        Return ""

    End Function

    Friend Function External() As Boolean

        Select Case My.Settings.SabPort
            Case 8090
                ''
            Case Else
                Return True
        End Select

        Select Case My.Settings.SabHost.ToLower.Trim
            Case "localhost", "localhost.", "127.0.0.1", "0.0.0.0"
                ''
            Case Else
                Return True
        End Select

        Return False

    End Function

    Private Sub SabCol_ProgressChanged(ByVal lVal As Integer) Handles SabCol.ProgressChanged

        RaiseEvent ProgressChanged(lVal)

    End Sub

    Private Function ProcessSab(ByVal xDataSlots As List(Of String()), ByVal xHistorySlots As List(Of String()), Optional ByRef zErr As String = "") As Boolean

        Static DidOnce As Boolean = False

        Try

            If SabStarted Then

                Dim FD As New Dictionary(Of String, String)

                For Each XZ() As String In xDataSlots

                    SabCol.Update(XZ(4), False, XZ(0), XZ(1), CInt(XZ(3)), XZ(5), XZ(2), CInt(XZ(6)), XZ(7), Single.Parse(XZ(8), System.Globalization.NumberFormatInfo.InvariantInfo))

                    If Not FD.ContainsKey(XZ(4).ToUpper) Then
                        FD.Add(XZ(4).ToUpper, XZ(4))
                    End If

                Next

                For Each XZ() As String In xHistorySlots

                    SabCol.Update(XZ(4), True, XZ(0), XZ(1), CInt(XZ(3)), XZ(5), XZ(2), CInt(XZ(6)), XZ(7), Single.Parse(XZ(8), System.Globalization.NumberFormatInfo.InvariantInfo))

                    If Not FD.ContainsKey(XZ(4).ToUpper) Then
                        FD.Add(XZ(4).ToUpper, XZ(4))
                    End If

                Next

                Dim DidRemove As Boolean = False

                Do
                    DidRemove = False

                    For Each Rex As String In SabCol.FD.Keys
                        If Not FD.ContainsKey(Rex) Then
                            SabCol.RemoveID(Rex)
                            DidRemove = True
                            Exit For
                        End If
                    Next

                Loop While DidRemove

                Return True


            Else
                zErr = "SAB niet gestart?"
            End If

            Return False

        Catch ex As Exception

            If Not DidOnce Then
                DidOnce = True
                Foutje("UpdateSab: " & ex.Message)
            End If

            Return False

        End Try

    End Function

    Private Sub SabUpdater_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles SabUpdater.DoWork

        Dim xDataSlots As New List(Of String())
        Dim xHistorySlots As New List(Of String())

        Try

            If GetSabQueue(xDataSlots, "") Then

                If GetSabHistory(xHistorySlots, "") Then

                    Dim xOut As New SabSlots

                    xOut.DataSlots = xDataSlots
                    xOut.HistorySlots = xHistorySlots

                    e.Result = xOut

                End If

            End If

        Catch
        End Try

    End Sub

    Private Sub SabUpdater_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles SabUpdater.RunWorkerCompleted

        If Not e.Result Is Nothing Then

            Dim zErr As String = ""
            Dim xOut As SabSlots = DirectCast(e.Result, SabSlots)

            Try
                If SabStarted Then
                    If Not SabUpdate Is Nothing Then

                        ProcessSab(xOut.DataSlots, xOut.HistorySlots, zErr)

                    End If
                End If
            Catch
            End Try

        End If

        SabUpdater = Nothing

    End Sub

End Class
