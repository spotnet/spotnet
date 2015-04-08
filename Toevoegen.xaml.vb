Imports System.Net
Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.ComponentModel
Imports System.Windows.Media.Imaging
Imports Spotnet.Spotnet
Imports System.Windows.Threading

Public Class Toevoegen

    Friend Sub New()

        InitializeComponent()

    End Sub

    Private Sub CatBox_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles CatBox.SelectionChanged

        Cat2Box.Items.Clear()

        SCatBox.Items.Clear()
        SCatBox.IsEnabled = False

        Label1.Content = TranslateCatDesc(CatBox.SelectedIndex + 1, "a0")

        Select Case CatBox.SelectedIndex

            Case 0

                Cat2Box.IsEnabled = True

                Cat2Box.Items.Add(NewItem("Films", 0))
                Cat2Box.Items.Add(NewItem("Series", 1))
                Cat2Box.Items.Add(NewItem("Boeken", 2))
                Cat2Box.Items.Add(NewItem("Erotiek", 3))

            Case 1
                Cat2Box.IsEnabled = True

                Cat2Box.Items.Add(NewItem("Album", 0))
                Cat2Box.Items.Add(NewItem("Liveset", 1))
                Cat2Box.Items.Add(NewItem("Podcast", 2))
                Cat2Box.Items.Add(NewItem("Audiobook", 3))

            Case Else

                Cat2Box.IsEnabled = False
                Cat2Box_SelectionChanged(sender, e)

        End Select

        UpdateBoxes()

    End Sub

    Private Function NewItem(ByVal sName As String, ByVal iTag As Integer) As ComboBoxItem

        Dim kl As New ComboBoxItem

        kl.Tag = iTag
        kl.Content = sName

        Return kl

    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button2.Click

        Me.Cursor = System.Windows.Input.Cursors.Wait
        Mouse.OverrideCursor = Cursors.Wait

        Button2.IsEnabled = False
        Me.IsEnabled = False

        Me.UpdateLayout()
        Me.Dispatcher.BeginInvoke(New Action(AddressOf Me.DoButton), DispatcherPriority.Background)

    End Sub

    Private Sub DoButton()

        Dim zErr As String = ""

        Try
            If DoPost(zErr) Then
                MsgBox("Je spot is succesvol toegevoegd." & vbCrLf & vbCrLf & "Het kan even duren voor hij in de database verschijnt.", MsgBoxStyle.Information, "Bedankt")
                My.Settings.Nickname = StripNonAlphaNumericCharacters(txtFrom.Text)
                My.Settings.Tagname = StripNonAlphaNumericCharacters(TxtTag.Text)
                My.Settings.Save()
                Dim xx As Spotnet.CloseableTabItem
                xx = CType(Me.Parent, Spotnet.CloseableTabItem)
                xx.CloseMe()
                Exit Sub
            End If
        Catch ex As Exception
            zErr = ex.Message
        End Try

        Foutje(zErr)

        Me.IsEnabled = True

        Me.Cursor = Nothing
        Mouse.OverrideCursor = Nothing

        Button2.IsEnabled = True

    End Sub

    Private Function CheckValues(ByRef zErr As String) As Boolean

        If CatBox.SelectedIndex < 0 Then zErr = "Selecteer een categorie!" : Return False
        If (Cat2Box.IsEnabled And Cat2Box.SelectedIndex < 0) Then zErr = "Selecteer een type!" : Return False
        If SCatBox.SelectedIndex < 0 Then zErr = "Selecteer een sub-categorie!" : Return False

        If Len(Trim(txtTitel.Text)) < 3 Then zErr = "Vul een titel in!" : Return False
        If Len(Trim(txtDesc.Text)) = 0 Then zErr = "Vul een beschrijving in!" : Return False
        If Len(Trim(txtDesc.Text)) < 50 Then zErr = "Beschrijving is te kort!" : Return False

        If Len(GetSubCats(CByte(CatBox.SelectedIndex), False)) = 0 Then zErr = "Selecteer minimaal 1 categorie!!" : Return False

        If Len(Trim(txtImage.Text)) = 0 Then zErr = ("Voeg eerst een een afbeelding toe!!") : Return False
        If Len(Trim(txtNZB.Text)) = 0 Then zErr = ("Voeg eerst een NZB bestand toe!") : Return False
        If Len(Trim(txtFrom.Text)) < 3 Then zErr = ("Afzender niet ingevuld!") : Return False

        Return True

    End Function

    Private Function DoPost(ByRef zErr As String) As Boolean

        Dim Rez() As Byte = Nothing

        Dim SizeX As Long = 0
        Dim SizeY As Long = 0

        If Not CheckValues(zErr) Then Return False

        If Not FileExists(Trim$(txtImage.Text)) Then
            If Not CheckPlaatje(AddHttp(Trim$(txtImage.Text)), Rez, SizeX, SizeY) Then zErr = ("Kan afbeelding niet vinden? Controleer of de URL correct is!") : Return False
        Else
            If Not CheckPlaatjeLocal(Trim$(txtImage.Text), Rez, SizeX, SizeY) Then zErr = ("Kan afbeelding niet toevoegen!") : Return False
        End If

        If Len(Trim(txtUrl.Text)) > 0 Then
            If Not CheckUrl(AddHttp(Trim$(txtUrl.Text))) Then zErr = ("Kan de website niet vinden? Controleer of de URL correct is!") : Return False
        End If

        Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)

        Return Spotlib.Spots.CreateSpot(UploadPhuse, My.Settings.HeaderGroup, txtTitel.Text, txtDesc.Text, CByte(CatBox.SelectedIndex + 1), GetSubCats(CByte(CatBox.SelectedIndex), True), AddHttp(txtUrl.Text.Trim), "nl", SizeX, SizeY, txtNZB.Text, txtFrom.Text, TxtTag.Text, My.Settings.NZBGroup, GetKey, CreateMsgID, Rez, GetAvatar, False, Ref.HeaderSettings(False, False), zErr)

    End Function

    Private Sub GetCats(ByVal hCat As Byte, ByVal hType As Integer)

        Dim Zx(2) As String
        Dim Sl As List(Of String)

        Zx(0) = "b"
        Zx(1) = "c"
        Zx(2) = "d"

        For ii = 0 To UBound(Zx)

            Sl = New List(Of String)

            Dim Catz As New Collection

            Select Case hCat
                Case 1

                    Select Case ii

                        Case 0

                            Select Case hType

                                Case 2

                                    Catz.Add(3)
                                    Catz.Add(10)

                                Case Else

                                    For i As Integer = 0 To 15
                                        If i <> 10 Then Catz.Add(i)
                                    Next

                            End Select

                        Case 1

                            Select Case hType

                                Case 2

                                    Catz.Add(2)
                                    Catz.Add(4)
                                    Catz.Add(12)
                                    Catz.Add(13)
                                    Catz.Add(14)
                                    Catz.Add(15)

                            End Select


                        Case 2

                            Select Case hType

                                Case 0, 1

                                    For i As Integer = 0 To 22
                                        Catz.Add(i)
                                    Next

                                    Catz.Add(27)
                                    Catz.Add(28)
                                    Catz.Add(29)

                                    Catz.Add(32)
                                    Catz.Add(33)
                                    Catz.Add(41)
                                    Catz.Add(50)
                                    Catz.Add(51)
                                    Catz.Add(54)

                                Case 2

                                    Catz.Add(1)
                                    Catz.Add(5)
                                    Catz.Add(7)
                                    Catz.Add(9)
                                    Catz.Add(15)
                                    Catz.Add(16)
                                    Catz.Add(17)
                                    Catz.Add(21)
                                    Catz.Add(30)
                                    Catz.Add(31)

                                    For i As Integer = 33 To 60
                                        Catz.Add(i)
                                    Next

                                Case 3

                                    Catz.Add(23)
                                    Catz.Add(24)
                                    Catz.Add(25)
                                    Catz.Add(26)

                                    For i As Integer = 72 To 90
                                        Catz.Add(i)
                                    Next

                            End Select

                    End Select

            End Select

            If Catz.Count = 0 Then
                For i As Integer = 0 To 100
                    Catz.Add(i)
                Next
            End If

            Dim zAdd As String = ""

            For Each i As Integer In Catz

                If hCat = 1 And hType = 2 And ii = 1 Then

                    zAdd = TranslateCat(5, Zx(ii) & i, True)

                Else

                    zAdd = TranslateCat(hCat, Zx(ii) & i, True)

                End If

                If Len(zAdd) > 0 Then Sl.Add(zAdd & vbTab & i)

            Next

            'Sl.Sort()

            For Each St As String In Sl

                Dim KL As New ListBoxItem
                Dim kl2 As New CheckBox
                Dim kl3 As String = Split(St, vbTab)(0)

                kl2.Content = kl3
                KL.Content = kl2
                KL.Tag = Val(Split(St, vbTab)(1))

                Select Case ii
                    Case 0
                        Cat1.Items.Add(KL)
                    Case 1
                        Cat2.Items.Add(KL)
                    Case 2
                        Cat3.Items.Add(KL)
                End Select

            Next

        Next

        If Cat1.Items.Count = 0 Then
            Cat1.Visibility = Windows.Visibility.Hidden
            CatLab1.Visibility = Windows.Visibility.Hidden
        Else
            Cat1.IsEnabled = True
            Cat1.Visibility = Windows.Visibility.Visible
            CatLab1.Content = TranslateCatDesc(hCat, "b0")
            CatLab1.Visibility = Windows.Visibility.Visible
        End If

        If Cat2.Items.Count = 0 Then
            Cat2.Visibility = Windows.Visibility.Hidden
            CatLab2.Visibility = Windows.Visibility.Hidden
        Else
            Cat2.IsEnabled = True
            Cat2.Visibility = Windows.Visibility.Visible
            CatLab2.Content = TranslateCatDesc(hCat, "c0")
            CatLab2.Visibility = Windows.Visibility.Visible
        End If

        If Cat3.Items.Count = 0 Then
            Cat3.Visibility = Windows.Visibility.Hidden
            CatLab3.Visibility = Windows.Visibility.Hidden
        Else
            Cat3.IsEnabled = True
            Cat3.Visibility = Windows.Visibility.Visible
            CatLab3.Content = TranslateCatDesc(hCat, "d0")
            CatLab3.Visibility = Windows.Visibility.Visible
        End If

    End Sub

    Private Sub UpdateBoxes()

        Cat1.Items.Clear()
        Cat1.IsEnabled = False

        Cat2.Items.Clear()
        Cat2.IsEnabled = False

        Cat3.Items.Clear()
        Cat3.IsEnabled = False

        If Not (CatBox.SelectedIndex < 0) Then

            If (Cat2Box.SelectedIndex >= 0) Or (Not Cat2Box.IsEnabled) Then

                GetCats(CByte(CatBox.SelectedIndex + 1), Cat2Box.SelectedIndex)

            End If

        End If

    End Sub

    Private Function GetSubCats(ByVal hCat As Byte, ByVal sIncludeHCat As Boolean) As String

        Dim KZ As CheckBox
        Dim GC As String = vbNullString

        Try

            If sIncludeHCat Then
                Dim CB As ComboBoxItem = CType(SCatBox.SelectedItem, ComboBoxItem)
                Dim aCat As Byte = CByte(CB.Tag)
                If aCat > 9 Then
                    GC += "A" & CStr(aCat)
                Else
                    GC += "A0" & CStr(aCat)
                End If
                If hCat = 0 Then

                    Dim CB2 As ComboBoxItem = CType(Cat2Box.SelectedItem, ComboBoxItem)

                    If Not CB2 Is Nothing Then

                        Dim bCat As Byte = CByte(CB2.Tag)

                        If bCat = 1 Then
                            GC += "B04D11"
                        End If

                        If bCat = 3 Then

                            Dim bFound As Boolean

                            For Each LL As ListBoxItem In Cat3.Items

                                KZ = CType(LL.Content, CheckBox)

                                If KZ.IsChecked Then
                                    Select Case CByte(LL.Tag)
                                        Case 23
                                            bFound = True
                                            GC += "D75"
                                        Case 24
                                            bFound = True
                                            GC += "D74"
                                        Case 25
                                            bFound = True
                                            GC += "D73"
                                        Case 26
                                            bFound = True
                                            GC += "D72"
                                    End Select
                                End If
                            Next

                            If Not bFound Then GC += "D23D75"

                        End If
                    End If
                End If
            End If

            For Each LL As ListBoxItem In Cat1.Items
                KZ = CType(LL.Content, CheckBox)
                If KZ.IsChecked Then
                    If Val(LL.Tag) > 9 Then
                        GC += "B" & CStr(LL.Tag)
                    Else
                        GC += "B0" & CStr(LL.Tag)
                    End If
                End If
            Next

            For Each LL As ListBoxItem In Cat2.Items
                KZ = CType(LL.Content, CheckBox)
                If KZ.IsChecked Then
                    If Val(LL.Tag) > 9 Then
                        GC += "C" & CStr(LL.Tag)
                    Else
                        GC += "C0" & CStr(LL.Tag)
                    End If
                End If
            Next

            For Each LL As ListBoxItem In Cat3.Items
                KZ = CType(LL.Content, CheckBox)
                If KZ.IsChecked Then
                    If Val(LL.Tag) > 9 Then
                        GC += "D" & CStr(LL.Tag)
                    Else
                        GC += "D0" & CStr(LL.Tag)
                    End If
                End If
            Next

            If sIncludeHCat Then
                If hCat < 2 Then
                    Dim CB2 As ComboBoxItem = CType(Cat2Box.SelectedItem, ComboBoxItem)
                    If Not CB2 Is Nothing Then GC += "Z0" & CStr(CByte(CB2.Tag))
                End If
            End If

            Return GC

        Catch ex As Exception
            Return ""
        End Try

    End Function

    Private Function GetHTMLPage(ByVal URL As String, Optional ByVal TimeoutSeconds As Integer = 5, Optional ByVal proxy As Integer = -1, Optional ByVal dAuth As Boolean = False) As String

        ' Retrieves the HTML from the specified URL,
        ' using a default timeout of 10 seconds

        Dim objRequest As Net.WebRequest
        Dim objResponse As Net.WebResponse
        Dim objStreamReceive As System.IO.Stream
        Dim objEncoding As System.Text.Encoding
        Dim objStreamRead As System.IO.StreamReader

        Try
            ' Setup our Web request

            objRequest = Net.WebRequest.Create(URL)

            objRequest.Proxy = Nothing
            objRequest.Timeout = TimeoutSeconds * 1000

            ' Retrieve data from request
            objResponse = objRequest.GetResponse
            objStreamReceive = objResponse.GetResponseStream
            objEncoding = System.Text.Encoding.GetEncoding("utf-8")
            objStreamRead = New System.IO.StreamReader(objStreamReceive, objEncoding)

            ' Set function return value
            GetHTMLPage = objStreamRead.ReadToEnd()

            ' Check if available, then close response
            If Not objResponse Is Nothing Then
                objResponse.Close()
            End If

        Catch ex As Exception
            ' Error occured grabbing data, simply return nothing
            Return "ERROR"

        End Try

    End Function

    Private Function CheckPlaatje(ByVal sUrl As String, ByRef Rez() As Byte, ByRef SizeX As Long, ByRef SizeY As Long) As Boolean

        Dim bFound As Boolean

        Try

            Dim fileReader As New WebClient()

            If Not HasHttp(sUrl) Then Return False

            Rez = fileReader.DownloadData(sUrl)

            For Each xx As String In fileReader.ResponseHeaders.Keys
                Select Case (fileReader.ResponseHeaders(xx))
                    Case "image/png", "image/gif", "image/jpeg", "image/bmp"
                        bFound = UBound(Rez) > 10
                End Select
            Next

            If Not bFound Then Return False

            Dim oMS As New MemoryStream(Rez)
            Dim PL2 As BitmapFrame

            PL2 = BitmapFrame.Create(oMS, BitmapCreateOptions.DelayCreation, BitmapCacheOption.None)

            If (PL2.PixelWidth) > 10 And PL2.PixelHeight > 10 Then
                SizeX = PL2.PixelWidth
                SizeY = PL2.PixelHeight
                oMS.Close()
                Return True
            End If

            oMS.Close()
            Return False

        Catch ex As HttpListenerException
            Return False
        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Function CheckPlaatjeLocal(ByVal sUrl As String, ByRef Rez() As Byte, ByRef SizeX As Long, ByRef SizeY As Long) As Boolean

        Try

            If Not FileExists(sUrl) Then Return False

            Dim oMS As New FileStream(sUrl, FileMode.Open, FileAccess.Read)
            Dim PL2 As BitmapFrame

            PL2 = BitmapFrame.Create(oMS, BitmapCreateOptions.DelayCreation, BitmapCacheOption.None)

            If (PL2.PixelWidth) > 10 And PL2.PixelHeight > 10 Then
                SizeX = PL2.PixelWidth
                SizeY = PL2.PixelHeight
                Dim oMS2 As New FileStream(sUrl, FileMode.Open, FileAccess.Read)
                ReDim Rez(CInt(oMS2.Length - 1))
                oMS2.Read(Rez, 0, CInt(oMS2.Length))
                oMS2.Close()
                Return True
            End If

            oMS.Close()
            Return False

        Catch ex As HttpListenerException
            Return False
        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Function CheckUrl(ByVal sUrl As String) As Boolean

        Try

            If Not HasHttp(sUrl) Then Return False

            Dim url As New System.Uri(sUrl)
            Dim req As System.Net.WebRequest
            req = System.Net.WebRequest.Create(url)

            req.Proxy = Nothing
            Dim resp As System.Net.WebResponse

            resp = req.GetResponse()
            Return (resp.ContentType.Substring(0, 4).ToLower = "text")

        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button1.Click

        Dim fdlg As Forms.OpenFileDialog = New Forms.OpenFileDialog()

        fdlg.Title = "NZB toevoegen"
        fdlg.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
        fdlg.Filter = "NZB Bestanden (*.nzb)|*.nzb"
        fdlg.FilterIndex = 1
        fdlg.RestoreDirectory = True
        fdlg.CheckFileExists = True
        fdlg.ShowReadOnly = False
        fdlg.DefaultExt = "nzb"
        fdlg.Multiselect = False

        If Len(My.Settings.LastFolder.Trim) > 0 Then
            fdlg.InitialDirectory = My.Settings.LastFolder
        Else
            fdlg.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
        End If

        If fdlg.ShowDialog() = Forms.DialogResult.OK Then
            txtNZB.Text = fdlg.FileName
            My.Settings.LastFolder = Path.GetDirectoryName(fdlg.FileName)
            My.Settings.Save()
        End If

    End Sub

    Private Sub TxtTag_PreviewTextInput(ByVal sender As Object, ByVal e As System.Windows.Input.TextCompositionEventArgs) Handles TxtTag.PreviewTextInput

        If (Not Char.IsLetterOrDigit(CChar(e.Text))) Then
            e.Handled = True
        End If

    End Sub

    Private Sub TxtTag_TextChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.TextChangedEventArgs) Handles TxtTag.TextChanged

        If TxtTag.Text <> StripNonAlphaNumericCharacters(TxtTag.Text) Then
            TxtTag.Text = StripNonAlphaNumericCharacters(TxtTag.Text)
            TxtTag.SelectionStart = TxtTag.Text.Length
        End If

    End Sub

    Private Sub txtFrom_PreviewTextInput(ByVal sender As Object, ByVal e As System.Windows.Input.TextCompositionEventArgs) Handles txtFrom.PreviewTextInput

        If (Not Char.IsLetterOrDigit(CChar(e.Text))) Then
            e.Handled = True
        End If

    End Sub

    Private Sub txtFrom_TextChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.TextChangedEventArgs) Handles txtFrom.TextChanged

        If txtFrom.Text <> StripNonAlphaNumericCharacters(txtFrom.Text) Then
            txtFrom.Text = StripNonAlphaNumericCharacters(txtFrom.Text)
            txtFrom.SelectionStart = txtFrom.Text.Length
        End If

    End Sub

    Private Sub Toevoegen_Initialized(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Initialized

        txtFrom.Text = StripNonAlphaNumericCharacters(My.Settings.Nickname)
        TxtTag.Text = StripNonAlphaNumericCharacters(My.Settings.Tagname)

    End Sub

    Private Sub Cat2Box_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles Cat2Box.SelectionChanged

        Dim i As Long

        SCatBox.Items.Clear()
        SCatBox.IsEnabled = True

        Label1.Content = TranslateCatDesc(CatBox.SelectedIndex + 1, "a0")

        Dim iStart As Integer = 0
        Dim iStop As Integer = 100

        Dim aByte = 0

        Dim CB2 As ComboBoxItem = CType(Cat2Box.SelectedItem, ComboBoxItem)
        If Not CB2 Is Nothing Then aByte = CByte(CB2.Tag)

        For i = iStart To iStop

            If CatBox.SelectedIndex = 0 Then
                If aByte = 2 Then
                    If i <> 5 Then Continue For
                Else
                    If i = 5 Then Continue For
                End If
            End If

            Dim sCat As String = TranslateCat(CatBox.SelectedIndex + 1, "a" & i, True)

            If Not sCat Is Nothing Then
                If Len(sCat) > 0 Then

                    Dim kl As New ComboBoxItem

                    kl.Tag = i
                    kl.Content = sCat

                    SCatBox.Items.Add(kl)

                End If
            End If

        Next

        UpdateBoxes()

    End Sub

    Private Sub Button3_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button3.Click

        Dim fdlg As Forms.OpenFileDialog = New Forms.OpenFileDialog()

        fdlg.Title = "Afbeeldingen toevoegen"
        fdlg.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
        fdlg.Filter = "Afbeeldingen|*.jpg;*.gif;*.png;*.bmp|JPG Bestand (*.jpg)|*.jpg|GIF Bestand (*.gif)|*.gif|PNG Bestand (*.png)|*.png|BMP Bestand (*.bmp)|*.bmp"
        fdlg.FilterIndex = 1
        fdlg.RestoreDirectory = True
        fdlg.CheckFileExists = True
        fdlg.ShowReadOnly = False
        fdlg.DefaultExt = "jpg"
        fdlg.Multiselect = False

        fdlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)

        If fdlg.ShowDialog() = Forms.DialogResult.OK Then
            txtImage.Text = fdlg.FileName
        End If

    End Sub

End Class
