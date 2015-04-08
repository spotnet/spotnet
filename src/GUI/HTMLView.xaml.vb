Imports mshtml
Imports System.IO
Imports System.Xml
Imports System.Drawing
Imports System.Data.Common
Imports System.Windows.Threading
Imports System.ComponentModel
Imports System.Drawing.Drawing2D

Imports Spotlib
Imports Spotnet.Spotnet

Public Class HTMLView

    Private bAllowNavi As Boolean

    Private LastTime As Date
    Private LastBody As String = ""
    Private DatabaseFile As String = ""

    Private CacheXoverID As Long = -1

    Private AskUnload As Boolean = False
    Private Const XGR As String = "Geen reacties"

    Private WithEvents _document As Forms.HtmlDocument
    Private WithEvents CommentStarter As BackgroundWorker
    Private WithEvents ImageStarter As BackgroundWorker
    Private WithEvents ExternalStarter As BackgroundWorker

    Private LastClick As Forms.HtmlElement
    Private CommentProgress As Forms.HtmlElement
    Private CommentsStatus As Forms.HtmlElement

    Private WithEvents AddButton As Forms.HtmlElement
    Private WithEvents SpotImage As Forms.HtmlElement
    Private WithEvents DownloadButton As Forms.HtmlElement

    Private WithEvents CommentLoader As Comments
    Private WithEvents CommentUpdater As Comments

    Private CommentIDCache As New HashSet(Of Long)
    Private UniqueCache As New Dictionary(Of String, String)
    Private CommentProgressCache As String
    Private FetchNewComments As Boolean = True

    Private MenuFrom As String
    Private MenuQuery As String
    Private MenuQueryName As String
    Private MenuModulus As String
    Private WithEvents SpotMenu As ContextMenu

    Private FetchedCache As New List(Of Long)
    Private SkipMessages As New HashSet(Of String)
    Private FetchedCacheHash As New HashSet(Of Long)

    Public Sub New()

        InitializeComponent()

    End Sub

    Friend Function GetWb(ByVal xAllowNavi As Boolean, ByVal dbFile As String) As System.Windows.Forms.WebBrowser

        DatabaseFile = dbFile
        bAllowNavi = xAllowNavi

        Try

            With Brows
                .ScriptErrorsSuppressed = bAllowNavi
                .WebBrowserShortcutsEnabled = bAllowNavi
                .IsWebBrowserContextMenuEnabled = True
            End With

        Catch
        End Try

        Return Brows

    End Function

    Private Sub brows_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Brows.Disposed

        Try
            Unload()
        Catch ex As Exception
            Foutje("Browser_Disposed: " & ex.Message)
        End Try

    End Sub

    Friend Sub FF()

        Try
            _document.Focus()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub brows_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles Brows.DocumentCompleted

        Try

            Static bOnlyOnce As Boolean = False

            If bAllowNavi Then

                _document = Brows.Document

                Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)
                Ref.UpdateTab(CType(Me.Parent, TabItem), _document.Title, _document.Url.AbsoluteUri)

            End If

            If bOnlyOnce Then Exit Sub
            If bAllowNavi Then Exit Sub
            If AskUnload Then Exit Sub

            bOnlyOnce = True

            Try

                AddButton = _document.GetElementById("AddComment")
                SpotImage = _document.GetElementById("SpotImage")
                DownloadButton = _document.GetElementById("DownloadButton")

                CommentProgress = _document.GetElementById("CommentsProgress")
                CommentProgressCache = CommentProgress.InnerHtml

            Catch
            End Try

            Dim xSpot As SpotEx = GetSpot()

            If Not SpotImage Is Nothing Then

                If Not ((Len(xSpot.Image) > 0) Or (Len(xSpot.ImageID) > 0)) Then

                    SpotImage.OuterHtml = ""

                Else

                    If Len(xSpot.ImageID) = 0 Then

                        If Len(xSpot.Web) > 0 Then
                            SpotImage.Style = "cursor:hand;" & SpotImage.Style
                        End If

                        SpotImage.SetAttribute("SRC", xSpot.Image)

                    Else

                        Me.Dispatcher.BeginInvoke(New Action(AddressOf Me.StartImage), DispatcherPriority.Background)
                        Exit Sub

                    End If

                End If

            End If

            xDoStart()

        Catch ex As Exception
            Foutje("DocumentCompleted: " & ex.Message)

        End Try

    End Sub

    Private Sub xDoStart()

        If AskUnload Then Exit Sub

        Me.Dispatcher.BeginInvoke(New Action(AddressOf Me.DoStart), DispatcherPriority.Background)

    End Sub

    Private Sub StartImage()

        If AskUnload Then Exit Sub

        ImageStarter = New BackgroundWorker
        ImageStarter.RunWorkerAsync(GetSpot.ImageID)

    End Sub

    Private Sub DoStart()

        Dim sErr As String = ""

        If AskUnload Then Exit Sub

        Try

            If (Not My.Settings.ShowComments) Then
                CommentsDone("", True)
                Exit Sub
            End If

            If StartUpdate(GetSpot.MessageID, sErr) Then Exit Sub

            CommentsDone(sErr)
            Exit Sub

        Catch ex As Exception

            CommentsDone("DoStart: " & ex.Message)
            Exit Sub

        End Try

    End Sub

    Private Function GetSpot() As SpotEx

        Try

            Dim DK As SpotInfo
            Dim zx2 As TabItem

            zx2 = CType(Me.Parent, TabItem)
            DK = CType(zx2.Tag, SpotInfo)

            Return DK.Spot

        Catch ex As Exception

            Foutje("GetSpot: " & ex.Message)
            Return Nothing

        End Try

    End Function

    Private Sub brows_DocumentTitleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Brows.DocumentTitleChanged

        Static OnlyOnce As Boolean = False

        If Not OnlyOnce Then
            OnlyOnce = True
            Try
                _document = Brows.Document
                Me.Dispatcher.BeginInvoke(New Action(AddressOf Me.FF), DispatcherPriority.Background)
            Catch
            End Try
        End If

        If bAllowNavi Then

            Dim zx2 As TabItem
            zx2 = CType(Me.Parent, TabItem)

            If TypeOf zx2.Tag Is UrlInfo Then
                If Not zx2 Is Nothing Then

                    _document = Brows.Document
                    If Not _document Is Nothing Then

                        If GetHeader(zx2.Header) <> _document.Title Then

                            Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)
                            If Len(_document.Title) > 0 Then zx2.Header = CreateHeader(_document.Title, "url.ico")
                            Ref.UpdateTab(CType(Me.Parent, TabItem), _document.Title, _document.Url.AbsoluteUri)

                            Exit Sub

                        End If
                    End If

                End If
            End If
        End If

    End Sub

    Private Sub brows_Navigating(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserNavigatingEventArgs) Handles Brows.Navigating

        Try

            If Not bAllowNavi Then

                Dim sUrl As String
                Dim sTitle As String

                sUrl = Trim$(e.Url.OriginalString)

                If sUrl.ToLower = "about:blank" Then
                    Exit Sub
                End If

                If sUrl.ToLower.StartsWith("res:") Then Exit Sub

                e.Cancel = True

                Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)

                If sUrl.ToLower.StartsWith("link:") Then
                    If sUrl.Substring(5).IndexOf("http://binsearch.info/") = 0 Then
                        sTitle = sUrl.Substring(5)
                        If sTitle.IndexOf("q=") > 0 Then sTitle = sTitle.Substring(sTitle.IndexOf("q=") + 2)
                        If sTitle.IndexOf("&") > 0 Then sTitle = sTitle.Substring(0, sTitle.IndexOf("&"))
                        sTitle = URLDecode(sTitle)
                    Else
                        sTitle = AddHttp(sUrl.Substring(5)).Split("/"c)(2)
                        Dim XS() As String = sTitle.Split("."c)
                        sTitle = XS(UBound(XS) - 1) & "." & XS(UBound(XS))
                    End If
                    If My.Settings.ExternalBrowser Then
                        LaunchBrowser(sUrl.Substring(5))
                    Else
                        Ref.OpenURL(sUrl.Substring(5), sTitle, Nothing)
                    End If
                    Exit Sub
                End If

                If sUrl.ToLower.StartsWith("query:") Then

                    Dim sError As String = Nothing

                    Ref.SearchFilter(URLDecode(sUrl.Substring(6).Split("_"c)(1)), URLDecode(sUrl.Substring(6).Split("_"c)(0)))

                    Exit Sub

                End If

                If sUrl.ToLower.StartsWith("menu:") Then
                    Dim sFrom As String = URLDecode(sUrl.Substring(5).Split("_"c)(1))
                    CreateMenu(sFrom, sUrl.Substring(5).Split("_"c)(0), "sender MATCH '" & sFrom.ToLower & "'", sFrom)
                    Exit Sub
                End If

                If sUrl.ToLower.StartsWith("spotnet:reload") Then
                    Dim sErr As String = ""
                    If Not StartUpdate(GetSpot.MessageID, sErr) Then
                        Foutje(sErr)
                    End If
                    Exit Sub
                End If

            End If

        Catch ex As Exception
            Foutje("Navigating: " & ex.Message)

        End Try

    End Sub

    Private Sub brows_NewWindow(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Brows.NewWindow

        e.Cancel = True

        Dim zUrl As String
        zUrl = CStr(sender.statustext)

        If zUrl Is Nothing Then Exit Sub
        If zUrl.Length < 6 Then Exit Sub

        If zUrl.Substring(0, 5).ToLower = "link:" Then
            LaunchBrowser(AddHttp(zUrl.Substring(5)))
        Else
            If HasHttp(zUrl) Then LaunchBrowser(zUrl)
        End If

    End Sub

    Private Sub DownloadButton_Click(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles DownloadButton.Click

        If Not DownloadButton.Enabled Then Exit Sub
        DisableDownload()

        Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)

        Ref.DoWait("Downloaden...")

        Me.Dispatcher.BeginInvoke(New Action(AddressOf Me.DoDownload), DispatcherPriority.Background)

    End Sub

    Private Sub DisableDownload()

        If Not DownloadButton Is Nothing Then
            DownloadButton.Enabled = False

            If DownloadButton.GetAttribute("src").ToLower.Contains("download.png") Or DownloadButton.GetAttribute("src").ToLower.Contains("download3.png") Then
                DownloadButton.SetAttribute("src", SettingsFolder() & "\Images\download2.png")
            End If

            DownloadButton.Style = Replace(DownloadButton.Style, "hand", "wait", , , CompareMethod.Text)
            DownloadButton.SetAttribute("title", "")

        End If

    End Sub

    Private Sub EnableDownload()

        If Not DownloadButton Is Nothing Then
            DownloadButton.Enabled = True

            If DownloadButton.GetAttribute("src").ToLower.Contains("download2.png") Or DownloadButton.GetAttribute("src").ToLower.Contains("download3.png") Then
                DownloadButton.SetAttribute("src", SettingsFolder() & "\Images\download.png")
            End If

            DownloadButton.Style = Replace(DownloadButton.Style, "wait", "hand", , , CompareMethod.Text)
            DownloadButton.SetAttribute("title", "Downloaden")
        End If

    End Sub

    Private Sub DoComment()

        Dim zErr As String = ""
        Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)

        Try

            Dim HashMsg As String = CreateMsgID(GetSpot.MessageID.Split("@"c)(0).Replace(".", "").Replace("<", ""))

            My.Settings.Nickname = StripNonAlphaNumericCharacters(CStr(_document.GetElementById("Nickname").DomElement.Value))
            My.Settings.Save()

            If Spotlib.Spots.CreateComment(UploadPhuse, CStr(_document.GetElementById("Nickname").DomElement.Value), CStr(_document.GetElementById("CommentBody").DomElement.Value), My.Settings.ReplyGroup, GetSpot.MessageID, GetSpot.Title, GetAvatar, GetKey, HashMsg, zErr) Then

                LastBody = CStr(_document.GetElementById("CommentBody").DomElement.Value)
                LastTime = Now

                _document.GetElementById("CommentBody").DomElement.Value = ""

                Dim PL As New Comment

                PL.Created = Now
                PL.From = My.Settings.Nickname
                PL.Body = LastBody
                PL.MessageID = MakeMsg(HashMsg, False)
                PL.User = New UserInfo

                If Not GetAvatar() Is Nothing Then
                    PL.User.Avatar = My.Settings.Avatar
                End If

                PL.User.Signature = PL.MessageID
                PL.User.Modulus = GetModulus()
                PL.User.ValidSignature = True

                NewComment(PL, True)
                SkipMessages.Add(MakeMsg(HashMsg, True))

                Ref.EndWait(True)
                EnableAdd()

                Exit Sub

            End If

        Catch ex As Exception
            zErr = ex.Message
        End Try

        Ref.EndWait(True)
        Foutje(zErr)
        EnableAdd()

    End Sub

    Private Sub DoDownload()

        Try

            Dim sFile As String = ""

            With GetSpot()

                Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)

                If .OldInfo Is Nothing Then

                    sFile = OpenNZB(.NZB, .Title)
                    Ref.EndWait(True)

                Else

                    sFile = FindNZB(HtmlDecode(.OldInfo.FileName), .OldInfo.Groups.Split("|"c)(0), .Title)
                    Ref.EndWait(True)

                End If

                If Len(sFile) > 0 Then

                    Select Case My.Settings.DownloadAction

                        Case 0, 1

                            If Ref.DownloadNZB(sFile, .Title) Then

                                Dim zx As TabControl
                                Dim zx2 As TabItem

                                zx2 = CType(Me.Parent, TabItem)
                                zx = CType(zx2.Parent, TabControl)
                                zx.SelectedIndex = 1

                            End If

                        Case 2

                            ExternalStarter = New BackgroundWorker
                            ExternalStarter.RunWorkerAsync(sFile)
                            Exit Sub

                        Case 3

                            Dim rFile As String = Ref.AskFile(.Title)

                            If Len(rFile) > 0 Then

                                My.Settings.LastFolder = System.IO.Path.GetDirectoryName(rFile)
                                My.Settings.Save()

                                System.IO.File.Copy(sFile, rFile, True)

                            End If

                    End Select

                End If

            End With

            EnableDownload()

        Catch ex As Exception

            Foutje(ex.Message)

        End Try

    End Sub

    Public Function CancelComments() As Boolean

        If Not CommentLoader Is Nothing Then
            CommentLoader.Cancel()
        End If

        If Not CommentUpdater Is Nothing Then
            CommentUpdater.Cancel()
        End If

        Return True

    End Function

    Public Function Unload() As Boolean

        If AskUnload Then Return True

        Try

            AskUnload = True
            CancelComments()

            CommentLoader = Nothing
            CommentUpdater = Nothing

            Brows.Stop()
            Brows.Dispose()
            WindowsFormsHost1.Dispose()

            Return True

        Catch ex As Exception

            Foutje("HTMLView_Unload: " & ex.Message)
            Return False

        End Try

    End Function

    Private Sub AddButton_Click(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles AddButton.Click

        If Not AddButton.Enabled Then Exit Sub

        If StripNonAlphaNumericCharacters(CStr(_document.GetElementById("CommentBody").DomElement.Value)).Trim.Length = 0 Then
            MsgBox("Je kunt geen leeg bericht plaatsen.", MsgBoxStyle.Information, "Fout")
            Exit Sub
        End If

        If StripNonAlphaNumericCharacters(CStr(_document.GetElementById("CommentBody").DomElement.Value)).ToLower = StripNonAlphaNumericCharacters(LastBody).ToLower Then
            MsgBox("Je kunt niet twee keer achter elkaar hetzelfde bericht plaatsen.", MsgBoxStyle.Information, "Fout")
            Exit Sub
        End If

        If DateDiff("s", LastTime, Now) < 10 Then
            MsgBox("Je moet " & 10 - DateDiff("s", LastTime, Now) & " seconden wachten tot je het volgende bericht kan plaatsen.", MsgBoxStyle.Information, "Fout")
            Exit Sub
        End If

        DisableAdd()

        Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)

        Ref.DoWait("Reactie plaatsen...")

        Me.Dispatcher.BeginInvoke(New Action(AddressOf Me.DoComment), DispatcherPriority.Background)

    End Sub

    Private Function StartUpdate(ByVal sMsgID As String, ByRef zErr As String) As Boolean

        If AskUnload Then
            zErr = "Exiting"
            Return False
        End If

        If (CommentProgress Is Nothing) Then
            zErr = "CommentProgress Is Nothing"
            Return False
        End If

        If CommentProgress.InnerHtml <> CommentProgressCache Then
            CommentProgress.InnerHtml = CommentProgressCache
        End If

        CommentsStatus = _document.GetElementById("CommentsStatus")

        ProgressChanged("Reacties laden...", -1)

        CommentStarter = New BackgroundWorker

        CommentStarter.WorkerReportsProgress = False
        CommentStarter.WorkerSupportsCancellation = False

        CommentStarter.RunWorkerAsync(sMsgID)

        Return True

    End Function

    Private Sub ShowComments(ByVal bCheckNewComments As Boolean)

        Try

            If AskUnload Then Exit Sub

            Dim NewList As New List(Of Long)

            For Each xArt As Long In FetchedCache
                If Not CommentIDCache.Contains(xArt) Then
                    NewList.Add(xArt)
                End If
            Next

            If NewList.Count > 0 Then
                FetchNewComments = bCheckNewComments
                Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)
                CommentLoader = Spots.GetComments(sModule.HeaderPhuse, NewList, Ref.CommentSettings(False, False))
            Else
                If bCheckNewComments Then
                    CheckNewComments()
                Else
                    CommentsDone("")
                End If
            End If

        Catch ex As Exception

            CommentsDone("ShowComments: " & ex.Message)

        End Try

    End Sub

    Private Sub CommentLoader_Completed(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs) Handles CommentLoader.Completed

        CommentLoader = Nothing

        If AskUnload Then Exit Sub

        Dim xReturn As CommentsCompletedEventArgs = CType(e, CommentsCompletedEventArgs)

        If xReturn Is Nothing Then
            CommentsDone("xReturn Is Nothing")
            Exit Sub
        End If

        If xReturn.Cancelled Or (Not xReturn.Error Is Nothing) Then
            If Len(xReturn.Error.Message) = 0 Then
                CommentsDone("Message Is Nothing")
            Else
                CommentsDone(xReturn.Error.Message)
            End If
            Exit Sub
        End If

        If Not FetchNewComments Then
            CommentsDone("")
            Exit Sub
        End If

        CheckNewComments()

    End Sub

    Private Sub CommentsDone(ByVal sError As String, Optional ByVal NotFetched As Boolean = False)

        Try

            If AskUnload Then Exit Sub

            Dim xReload As String = "<p><A onfocus='this.blur()' HREF='spotnet:reload'><IMG id='reload' onfocus='this.blur()' title='Vernieuwen' style='border: 0px; cursor:hand; width: 32px; height:32px;' SRC=" & Chr(34) & SettingsFolder() & "\Images\refresh.png" & Chr(34) & "></A>"

            If Not CommentProgress Is Nothing Then
                If Len(sError) > 0 Then
                    CommentProgress.InnerHtml = "<center>" & HtmlEncode(sError) & "<br></center>" & xReload
                Else
                    If CommentIDCache.Count = 0 Then
                        If Not NotFetched Then
                            CommentProgress.InnerHtml = "<center>" & "Geen reacties gevonden" & "<br></center>" & xReload
                        Else
                            CommentProgress.InnerHtml = "<center>" & "Reacties niet opgehaald" & "<br></center>" & xReload
                        End If
                    Else
                        CommentProgress.InnerHtml = xReload
                    End If
                End If
            End If

        Catch ex As Exception

            Foutje("CommentsDone: " & ex.Message)

        End Try

    End Sub

    Private Sub DisableAdd()

        If Not AddButton Is Nothing Then

            AddButton.Enabled = False

            If AddButton.GetAttribute("src").ToLower.Contains("reply.png") Then
                AddButton.SetAttribute("src", SettingsFolder() & "\Images\reply2.png")
            End If

            AddButton.Style = Replace(AddButton.Style, "hand", "wait", , , CompareMethod.Text)
            AddButton.SetAttribute("title", "")
        End If

    End Sub

    Private Sub EnableAdd()

        If Not AddButton Is Nothing Then
            If AddButton.GetAttribute("src").ToLower.Contains("reply2.png") Then
                AddButton.SetAttribute("src", SettingsFolder() & "\Images\reply.png")
            End If

            AddButton.Style = Replace(AddButton.Style, "wait", "hand", , , CompareMethod.Text)
            AddButton.SetAttribute("title", "Verzenden")
            AddButton.Enabled = True
        End If

    End Sub

    Private Sub NewComment(ByRef z As Comment, ByVal bVirtual As Boolean)

        If AskUnload Then Exit Sub
        If Len(z.MessageID) = 0 Then Exit Sub
        If CommentIDCache.Contains(z.Article) Then Exit Sub

        If BlackList.Contains(z.User.Modulus) Then
            CommentIDCache.Add(z.Article)
            Exit Sub
        End If

        If SkipMessages.Contains(MakeMsg(z.MessageID, True)) Then
            Exit Sub
        End If

        Try

            Dim sClass As String = "comment"

            Dim GS As SpotEx = GetSpot()

            Dim sTooltip As String = ""

            If Not UniqueCache.ContainsKey(GS.Poster.ToUpper) Then
                UniqueCache.Add(GS.Poster.ToUpper, GS.User.Modulus)
            End If

            z.From = StripNonAlphaNumericCharacters(z.From)

            If bVirtual Then

                sTooltip = HtmlEncode(MakeUnique(GetModulus))

            Else

                If (Len(z.User.Modulus) = 0) Or (Not z.User.ValidSignature) Then

                    sTooltip = "Onbekend"

                    If Len(z.User.Organisation) > 0 Then sTooltip += vbCrLf & HtmlEncode(z.User.Organisation)
                    If Len(z.User.Trace) > 3 Then sTooltip += vbCrLf & HtmlEncode(z.User.Trace)

                Else

                    If Not UniqueCache.ContainsKey(z.From.ToUpper) Then
                        UniqueCache.Add(z.From.ToUpper, z.User.Modulus)
                    End If

                    sTooltip = HtmlEncode(MakeUnique(z.User.Modulus))
                    If Len(z.User.Organisation) > 0 Then sTooltip += vbCrLf & HtmlEncode(z.User.Organisation)

                    If WhiteList.Contains(z.User.Modulus) Then

                        sClass = "trusted"

                    Else

                        If (z.User.Modulus = GS.User.Modulus) Then

                            sClass = "author"

                        Else

                            If Len(z.User.Trace) > 3 Then sTooltip += vbCrLf & HtmlEncode(z.User.Trace)

                            If UniqueCache.Item(z.From.ToUpper) <> z.User.Modulus Then

                                z.From = Trim(z.From) & " (2)"

                            End If

                        End If

                    End If

                End If

            End If

            Dim ReactDiv As Forms.HtmlElement
            ReactDiv = _document.GetElementById("Comments")

            If Not ReactDiv Is Nothing Then

                Dim tComment As Forms.HtmlElement = _document.CreateElement("SPAN")
                tComment.InnerHtml = SpotParser.ParseComment(z, sClass, sTooltip)
                ReactDiv.InsertAdjacentElement(Forms.HtmlElementInsertionOrientation.BeforeEnd, tComment)
                If ReactDiv.DomElement.Style.visibility <> "" Then ReactDiv.DomElement.Style.visibility = ""
                CommentIDCache.Add(z.Article)

            End If

        Catch ex As Exception

            Foutje("NewComment: " & ex.Message)

        End Try

    End Sub

    Private Sub CommentLoader_NewComment(ByVal e As SpotnetNewCommentEventArgs) Handles CommentLoader.NewComment

        If Not AskUnload Then NewComment(e.cComment, False)

    End Sub

    Private Sub CommentLoader_ProgressChanged(ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles CommentLoader.ProgressChanged

        If AskUnload Then Exit Sub

        Dim zx As SpotnetProgressChangedEventArgs = CType(e, SpotnetProgressChangedEventArgs)
        ProgressChanged(zx.ProgressMessage, zx.ProgressPercentage)

    End Sub

    Public Sub ProgressChanged(ByVal sMessage As String, ByVal sValue As Long)

        Try

            If AskUnload Then Exit Sub

            Dim sText As String = ""

            If CommentsStatus Is Nothing Then Exit Sub

            If (Len(sMessage) > 0) Then

                sText = sMessage

            Else

                If sValue > 100 Then sValue = 100

                If sValue > 0 Then
                    sText = "Reacties laden (" & CStr(sValue) & "%)"
                Else
                    sText = "Reacties laden..."
                End If

            End If

            If CommentsStatus.InnerText <> sText Then CommentsStatus.InnerText = sText

        Catch ex As Exception

        End Try

    End Sub

    Private Sub _document_Click(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles _document.Click

        LastClick = _document.GetElementFromPoint(e.MousePosition)

        DoClear()

    End Sub

    Private Sub _document_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles _document.MouseMove

        If e.MouseButtonsPressed = Forms.MouseButtons.Left Then DoClear()

    End Sub

    Private Sub DoClear()

        If bAllowNavi Then Exit Sub

        Try

            With _document.ActiveElement
                Select Case .TagName.ToUpper

                    Case "TEXTAREA", "INPUT"

                        Exit Sub

                End Select

            End With

            _document.DomDocument.selection.empty()

        Catch

        End Try

    End Sub

    Private Sub _document_ContextMenuShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles _document.ContextMenuShowing

        If bAllowNavi Then Exit Sub

        Try
            With _document.ActiveElement
                Select Case .TagName.ToUpper

                    Case "TEXTAREA", "INPUT"
                        Exit Sub

                    Case Else
                        e.BubbleEvent = False
                        e.ReturnValue = False

                End Select

            End With
        Catch
            e.BubbleEvent = False
            e.ReturnValue = False
        End Try

        Dim zx2 As TabItem
        Dim zx3 As CloseableTabItem

        Try
            zx2 = CType(Me.Parent, TabItem)
            zx3 = CType(zx2, CloseableTabItem)
            zx3.CloseMe()
            Exit Sub
        Catch
        End Try

    End Sub

    Private Sub _document_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles _document.MouseDown

        DoClear()

    End Sub

    Private Sub _document_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles _document.MouseLeave

        DoClear()

    End Sub

    Private Sub _document_MouseOver(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles _document.MouseOver

        DoClear()

    End Sub

    Private Sub _document_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles _document.MouseUp

        DoClear()

    End Sub

    Private Sub brows_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles Brows.PreviewKeyDown

        If e.KeyCode = 46 Then
            Try
                With _document.ActiveElement
                    Select Case .TagName.ToUpper

                        Case "TEXTAREA", "INPUT"
                            _document.DomDocument.selection.clear()

                        Case Else

                    End Select

                End With
            Catch
            End Try
        End If

    End Sub

    Private Sub DownloadButton_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles DownloadButton.MouseEnter

        If DownloadButton.GetAttribute("src").ToLower.Contains("download.png") Then
            DownloadButton.SetAttribute("src", SettingsFolder() & "\Images\download3.png")
        End If

    End Sub

    Private Sub DownloadButton_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles DownloadButton.MouseLeave

        If DownloadButton.GetAttribute("src").ToLower.Contains("download3.png") Then
            DownloadButton.SetAttribute("src", SettingsFolder() & "\Images\download.png")
        End If

    End Sub

    Private Sub CommentUpdater_Completed(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs) Handles CommentUpdater.Completed

        Dim sErr As String = ""
        Dim xMessageID As String = ""

        CommentUpdater = Nothing
        If AskUnload Then Exit Sub

        Dim xReturn As CommentsCompletedEventArgs = CType(e, CommentsCompletedEventArgs)

        If xReturn Is Nothing Then
            CommentsDone("xReturn Is Nothing")
            Exit Sub
        End If

        If xReturn.Cancelled Or (Not xReturn.Error Is Nothing) Then
            If Len(xReturn.Error.Message) = 0 Then
                CommentsDone("Message Is Nothing")
            Else
                CommentsDone(xReturn.Error.Message)
            End If
            Exit Sub
        End If

        If xReturn.Comments Is Nothing Then
            CommentsDone("xReturn.Comments Is Nothing")
        End If

        If xReturn.Comments.Count = 0 Then
            CommentsDone("")
            Exit Sub
        End If

        xMessageID = MakeMsg(GetSpot.MessageID, False)

        Dim bDidSome As Boolean = False

        For Each xComment As Comment In xReturn.Comments

            If xComment.MessageID = xMessageID Then

                If Not FetchedCacheHash.Contains(xComment.Article) Then
                    bDidSome = True
                    FetchedCache.Add(xComment.Article)
                    FetchedCacheHash.Add(xComment.Article)
                End If

            End If

        Next

        CacheXoverID = xReturn.Comments.Item(xReturn.Comments.Count - 1).Article

        If Not bDidSome Then
            CommentsDone("")
            Exit Sub
        End If

        ShowComments(False)

    End Sub

    Private Sub CommentUpdater_ProgressChanged(ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles CommentUpdater.ProgressChanged

        If AskUnload Then Exit Sub

        Dim zx As SpotnetProgressChangedEventArgs = CType(e, SpotnetProgressChangedEventArgs)

        If zx.ProgressMessage.Contains("gevonden") Then
            ProgressChanged("Reacties bijwerken (" & zx.ProgressPercentage & "%)", zx.ProgressPercentage)
        Else
            ProgressChanged(zx.ProgressMessage, zx.ProgressPercentage)
        End If

    End Sub

    Public Function OpenNZB(ByVal sLoc As String, ByVal sTitle As String) As String

        Dim zErr As String = ""
        Dim TheParts As New List(Of String)
        Dim sParts() As String = Split(sLoc, " ")

        For i As Integer = 0 To UBound(sParts)
            TheParts.Add(MakeMsg(sParts(i)))
        Next

        If TheParts.Count = 0 Then
            zErr = "Geen segmenten."
            GoTo Failz
        End If

        Dim zxOut As String = ""

        If Not Spots.GetNZB(DownloadPhuse, My.Settings.NZBGroup, TheParts, zxOut, zErr) Then GoTo Failz

        Dim m_xmld As XmlDocument

        Try

            m_xmld = New XmlDocument()
            m_xmld.XmlResolver = Nothing
            m_xmld.LoadXml(zxOut)

        Catch
            zErr = "Fout tijdens het parsen."
            GoTo Failz
        End Try

        Dim TmpFile As String
        Dim objReader As StreamWriter

        Dim sFile As String = Trim(MakeFilename(sTitle))

        If sFile.Length > 0 Then
            TmpFile = (System.IO.Path.GetTempPath & sFile & ".nzb")
        Else
            TmpFile = (System.IO.Path.GetTempFileName & ".nzb")
        End If

        Try

            objReader = New StreamWriter(TmpFile, False, LatinEnc)
            objReader.Write(zxOut)
            objReader.Close()

        Catch Ex As Exception
            zErr = "Fout tijdens het schrijven."
            GoTo Failz
        End Try

        Return TmpFile

Failz:
        Foutje(zErr)
        Return vbNullString

    End Function

    Public Function OpenImage(ByVal sLoc As String) As String

        Dim zErr As String = ""
        Dim TheParts As New List(Of String)
        Dim sParts() As String = Split(sLoc, " ")

        For i As Integer = 0 To UBound(sParts)
            TheParts.Add(MakeMsg(sParts(i), True))
        Next

        If TheParts.Count = 0 Then
            zErr = "Geen segmenten."
            GoTo Failz
        End If

        Dim zxOut() As Byte = Nothing
        If Not Spots.GetImage(DownloadPhuse, My.Settings.NZBGroup, TheParts, zxOut, zErr) Then Return vbNullString

        Dim TmpFile As String = (System.IO.Path.GetTempFileName)
        Dim objReader As StreamWriter

        Try

            objReader = New StreamWriter(TmpFile, False, LatinEnc)
            Dim t As New BinaryWriter(objReader.BaseStream, LatinEnc)

            t.Write(zxOut)
            objReader.Close()

        Catch Ex As Exception
            zErr = "Fout tijdens het schrijven: " & Ex.Message
            GoTo Failz
        End Try

        Return TmpFile

Failz:
        Foutje(zErr)
        Return vbNullString

    End Function

    Private Function GetComments(ByVal TheDB As SqlDB, ByVal xMsg As String, ByRef xErr As String) As List(Of Long)

        Dim lFnd As Long
        Dim sErr As String = ""
        Dim sMessageID As String = ""
        Dim DR As DbDataReader = Nothing

        Try

            sMessageID = MakeMsg(xMsg, False)
            Dim AtPos As Integer = sMessageID.IndexOf("@")
            sMessageID = sMessageID.Substring(0, AtPos)

            DR = TheDB.ExecuteReader("SELECT docid FROM comments WHERE spot MATCH '" & sMessageID.Replace("'"c, "") & "' ORDER BY docid ASC", sErr)

            If DR Is Nothing Then Throw New Exception("Datareader not available: " & sErr)

            With DR
                While .Read

                    If IsDBNull(.Item(0)) Then Continue While

                    lFnd = CLng(.Item(0))

                    If lFnd > 0 Then

                        If Not FetchedCacheHash.Contains(lFnd) Then
                            FetchedCache.Add(lFnd)
                            FetchedCacheHash.Add(lFnd)
                        End If

                    End If

                End While
            End With

            DR.Close()
            TheDB.Close()

            Return FetchedCache

        Catch ex As Exception

            Try
                DR.Close()
            Catch
            End Try

            Try
                TheDB.Close()
            Catch
            End Try

            xErr = "GetComments: " & ex.Message
            Return Nothing

        End Try

    End Function

    Private Sub CommentStarter_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles CommentStarter.DoWork

        Try

            e.Result = Nothing
            If AskUnload Then Exit Sub

            Dim zErr As String = ""
            Dim Db As SqlDB = New SqlDB

            If Not Db.Connect(DatabaseFile, True) Then
                Throw New Exception("Db.Connect")
            End If

            If Not Db.ExecuteNonQuery("PRAGMA temp_store = MEMORY;", "") = 0 Then Throw New Exception("PRAGMA temp_store")
            If Not Db.ExecuteNonQuery("PRAGMA cache_size = " & CStr(My.Settings.DatabaseCache), "") = 0 Then Throw New Exception("PRAGMA cache_size")

            Dim TheList As List(Of Long) = GetComments(Db, CType(e.Argument, String), zErr)
            If TheList Is Nothing Then Throw New Exception(zErr)

            e.Result = TheList

        Catch ex As Exception

            Foutje("CommentStarter: " & ex.Message)

        End Try

    End Sub

    Private Sub CommentStarter_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles CommentStarter.RunWorkerCompleted

        Try

            CommentStarter = Nothing
            If AskUnload Then Exit Sub

            If Not e.Result Is Nothing Then

                ShowComments(True)
                Exit Sub

            End If

            CommentsDone("") ' Is al een messagebox geweest 

        Catch ex As Exception

            CommentsDone("CommentStarter_Complete: " & ex.Message)

        End Try

    End Sub

    Private Sub CheckNewComments()

        Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)
        Dim cSet As NNTPSettings = Ref.CommentSettings(True, False)

        If CacheXoverID > cSet.Position Then
            cSet.Position = CacheXoverID
        End If

        CommentUpdater = Spots.FindComments(HeaderPhuse, cSet)

    End Sub

    Private Sub ImageStarter_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles ImageStarter.DoWork

        e.Result = Nothing

        If AskUnload Then Exit Sub

        Try
            e.Result = OpenImage(CType(e.Argument, String))
        Catch
        End Try

    End Sub

    Private Sub ImageStarter_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles ImageStarter.RunWorkerCompleted

        ImageStarter = Nothing

        If AskUnload Then Exit Sub

        Try

            Dim sImg As String = CType(e.Result, String)

            If Not sImg Is Nothing Then
                If sImg.Length > 0 Then

                    If Len(GetSpot.Web) > 0 Then
                        SpotImage.Style = "cursor:hand;" & SpotImage.Style
                    End If

                    SpotImage.SetAttribute("SRC", sImg)

                Else
                    SpotImage.OuterHtml = ""
                End If
            Else
                SpotImage.OuterHtml = ""
            End If

        Catch ex As Exception

            Foutje("ImageStarter: " & ex.Message)
            Exit Sub

        End Try

        DoStart()

    End Sub

    Private Sub ExternalStarter_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles ExternalStarter.DoWork

        Dim p As New System.Diagnostics.Process
        Dim s As New System.Diagnostics.ProcessStartInfo(CType(e.Argument, String))

        Try

            s.UseShellExecute = True
            s.WindowStyle = ProcessWindowStyle.Normal
            p.StartInfo = s
            p.Start()

        Catch ex As Exception
        End Try

    End Sub

    Private Sub ExternalStarter_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles ExternalStarter.RunWorkerCompleted

        ExternalStarter = Nothing

        If AskUnload Then Exit Sub

        Try

        Catch
            EnableDownload()
        End Try

    End Sub

    Private Sub SpotImage_Click(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles SpotImage.Click

        If Not SpotImage.Style Is Nothing Then
            If SpotImage.Style.Contains("hand") Then Brows.Navigate("link:" & SafeHref(GetSpot.Web))
        End If

    End Sub

    Private Sub SpotImage_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles SpotImage.MouseLeave

        Dim SI As IHTMLElement = CType(SpotImage.DomElement, IHTMLElement)
        Dim SI2 As IHTMLImgElement = CType(SpotImage.DomElement, IHTMLImgElement)

        If (SI2.width > 350) Then SI.style.width = "350px"

    End Sub

    Private Sub SpotImage_MouseOver(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles SpotImage.MouseOver

        Dim SI As IHTMLElement = CType(SpotImage.DomElement, IHTMLElement)
        SI.style.width = ""

    End Sub

    Private Sub CreateMenu(ByVal sFrom As String, ByVal sModulus As String, ByVal sQuery As String, ByVal sQueryName As String)

        Dim tk As MenuItem
        SpotMenu = New ContextMenu

        Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)

        SpotMenu.FontFamily = Ref.FontFamily
        SpotMenu.FontSize = Ref.FontSize - 1
        SpotMenu.FontStyle = Ref.FontStyle

        MenuQuery = sQuery
        MenuQueryName = sQueryName

        If Len(sFrom) > 0 Then
            MenuFrom = sFrom
            MenuModulus = sModulus
        End If

        If (sModulus = GetModulus()) Then

            tk = New MenuItem
            tk.Tag = "ava"

            tk.IsEnabled = True
            tk.Header = "Avatar wijzigen"

            tk.Icon = GetIcon("settings")

            SpotMenu.Items.Add(tk)

        Else

            tk = New MenuItem

            tk.Header = "Zoeken"
            tk.Icon = GetIcon("search")
            tk.IsEnabled = True

            If tk.IsEnabled Then
                tk.Icon.Opacity = 1
            Else
                tk.Icon.Opacity = 0.5
            End If

            tk.Tag = "search"

            SpotMenu.Items.Add(tk)

        End If

        If Len(sFrom) > 0 Then

            SpotMenu.Items.Add(New Separator)

            tk = New MenuItem

            tk.Tag = "fav"

            tk.IsEnabled = (Len(sModulus) > 0) And (Not BlackList.Contains(sModulus))

            If WhiteList.Contains(sModulus) Then
                tk.Header = "Verwijderen van witte lijst"
            Else
                tk.Header = "Toevoegen aan witte lijst"
            End If

            tk.Icon = GetIcon("favorite")

            If tk.IsEnabled Then
                tk.Icon.Opacity = 1
            Else
                tk.Icon.Opacity = 0.5
            End If

            SpotMenu.Items.Add(tk)

            tk = New MenuItem
            tk.Tag = "black"

            tk.IsEnabled = (Len(sModulus) > 0) And (Not WhiteList.Contains(sModulus)) And (sModulus <> GetModulus())

            If BlackList.Contains(sModulus) Then
                tk.Header = "Verwijderen van zwarte lijst"
            Else
                tk.Header = "Toevoegen aan zwarte lijst"
            End If

            tk.Icon = GetIcon("trash")

            If tk.IsEnabled Then
                tk.Icon.Opacity = 1
            Else
                tk.Icon.Opacity = 0.5
            End If

            SpotMenu.Items.Add(tk)

        End If

        SpotMenu.IsOpen = True

    End Sub

    Private Function IsAuthor(ByVal xElement As Forms.HtmlElement) As Boolean

        Dim Se As Forms.HtmlElement = xElement

        If Se Is Nothing Then Return False

        If Not Se.Id Is Nothing Then
            If Se.Id.ToLower = "comments" Then Return False
        End If

        Do While Not Se.Parent Is Nothing
            Se = Se.Parent
            If Se.Id Is Nothing Then Continue Do
            If Se.Id.ToLower = "comments" Then Return False
        Loop

        Return True

    End Function

    Private Function FindHref(ByVal xElement As Forms.HtmlElement) As Forms.HtmlElement

        Dim Se As Forms.HtmlElement = xElement

        If Se Is Nothing Then Return Nothing

        If Not Se.TagName Is Nothing Then
            If Se.TagName.ToLower = "a" Then Return Se
        End If

        Do While Not Se.Parent Is Nothing
            Se = Se.Parent
            If Se.TagName Is Nothing Then Continue Do
            If Se.TagName.ToLower = "a" Then Return Se
        Loop

        Return Nothing

    End Function

    Private Sub SpotMenu_PreviewMouseUp(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles SpotMenu.PreviewMouseUp

        Try

            If e Is Nothing Then Exit Sub
            If e.Source Is Nothing Then Exit Sub
            If e.Source.Tag Is Nothing Then Exit Sub

            Select Case CStr(e.Source.Tag).ToLower

                Case "fav"

                    If BlackList.Contains(MenuModulus) Then Exit Sub
                    If Len(MenuModulus) = 0 Then Exit Sub

                    If WhiteList.Contains(MenuModulus) Then

                        RemoveWhite(MenuModulus)

                        LastClick = FindHref(LastClick)

                        If Not LastClick Is Nothing Then

                            If Not IsAuthor(LastClick) Then
                                If MenuModulus = GetSpot.User.Modulus Then
                                    LastClick.SetAttribute("className", "author")
                                Else
                                    LastClick.SetAttribute("className", "comment")
                                End If
                            Else
                                LastClick.SetAttribute("className", "from")
                            End If

                        End If

                    Else

                        AddWhite(StripNonAlphaNumericCharacters(MenuFrom), MenuModulus)

                        LastClick = FindHref(LastClick)

                        If Not LastClick Is Nothing Then
                            LastClick.SetAttribute("className", "trusted")
                        End If

                    End If

                    Exit Sub

                Case "black"

                    If WhiteList.Contains(MenuModulus) Then Exit Sub
                    If Len(MenuModulus) = 0 Then Exit Sub

                    If BlackList.Contains(MenuModulus) Then

                        RemoveBlack(MenuModulus)

                        LastClick = FindHref(LastClick)

                        If Not LastClick Is Nothing Then

                            If Not IsAuthor(LastClick) Then
                                If MenuModulus = GetSpot.User.Modulus Then
                                    LastClick.SetAttribute("className", "author")
                                Else
                                    LastClick.SetAttribute("className", "comment")
                                End If
                            Else
                                LastClick.SetAttribute("className", "from")
                            End If

                        End If

                        ShowOnce("Je zult weer spots/reacties van deze afzender gaan ontvangen.", "Zwarte lijst")

                    Else

                        AddBlack(StripNonAlphaNumericCharacters(MenuFrom), MenuModulus)

                        LastClick = FindHref(LastClick)

                        If Not LastClick Is Nothing Then
                            LastClick.SetAttribute("className", "untrusted")

                        End If

                        ShowOnce("Je zult geen spots/reacties van deze afzender meer ontvangen.", "Zwarte lijst")

                    End If


                Case "search"

                    Dim sError As String = Nothing
                    Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)

                    Ref.SearchFilter(MenuQuery, MenuQueryName)

                Case "save"

                    Dim sError As String = Nothing
                    Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)

                    If Not Ref.SaveFilter(MenuQuery, MenuQueryName, sError) Then
                        Foutje(sError, "Fout")
                    End If

                Case "ava"

                    SelectAvatar()

            End Select

        Catch ex As Exception

            Foutje("SpotMenu: " & ex.Message)

        End Try

    End Sub

    Private Sub SelectAvatar()

        Dim fdlg As Forms.OpenFileDialog = New Forms.OpenFileDialog()

        fdlg.Title = "Avatar wijzigen"
        fdlg.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
        fdlg.Filter = "Afbeeldingen|*.gif;*.png;|GIF Bestand (*.gif)|*.gif|PNG Bestand (*.png)|*.png"
        fdlg.FilterIndex = 1
        fdlg.RestoreDirectory = True
        fdlg.CheckFileExists = True
        fdlg.ShowReadOnly = False
        fdlg.DefaultExt = "gif"
        fdlg.Multiselect = False

        fdlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)

        If fdlg.ShowDialog() = Forms.DialogResult.OK Then

            Dim bNew() As Byte

            If FileLen(fdlg.FileName) > 4000 Then

                bNew = CreateAvatar(fdlg.FileName)

            Else

                Dim zR As New FileStream(fdlg.FileName, FileMode.Open, FileAccess.Read)
                ReDim bNew(CInt(zR.Length - 1))
                zR.Read(bNew, 0, CInt(zR.Length))

            End If

            If bNew.Length > 4000 Then

                Foutje("Je avatar mag maximaal 4000 bytes groot zijn.", "Avatar wijzigen")
                Exit Sub

            End If

            My.Settings.Avatar = Convert.ToBase64String(bNew)
            My.Settings.Save()

        End If

    End Sub

    Private Function CreateAvatar(ByVal sFile As String) As Byte()

        Dim Img As Image = Image.FromFile(sFile)
        Dim Thumb As Image = New Bitmap(32, 32)
        Dim Graphic As Graphics = Graphics.FromImage(Thumb)

        Graphic.InterpolationMode = InterpolationMode.HighQualityBicubic
        Graphic.SmoothingMode = SmoothingMode.HighQuality
        Graphic.PixelOffsetMode = PixelOffsetMode.HighQuality
        Graphic.CompositingQuality = CompositingQuality.HighQuality

        Graphic.DrawImage(Img, 0, 0, 32, 32)

        Dim oMs As New MemoryStream()
        Thumb.Save(oMs, System.Drawing.Imaging.ImageFormat.Png)

        Dim zOut(CInt(oMs.Length - 1)) As Byte

        oMs.Position = 0
        oMs.Read(zOut, 0, CInt(oMs.Length))

        Return zOut

    End Function

End Class
