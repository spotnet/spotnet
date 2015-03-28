Imports System.IO
Imports System.Xml
Imports System.Net
Imports System.Text
Imports System.Threading
Imports System.ComponentModel
Imports System.Windows.Interop
Imports System.Windows.Threading
Imports System.Windows.Controls.Primitives
Imports System.Windows.Media.Effects
Imports System.Security.Cryptography
Imports DataVirtualization

Namespace Spotnet

    Public Class MainWindow

        Inherits Window

        Private LastSel As String
        Private TabDB As New cTabs
        Private FilterDB As New cFilter
        Private HistoryDB As New cHistory

        Private ProgX As Long = -1
        Private ProgXM As String = ""

        Private LastTab As Integer = 0
        Private SuggestList As New List(Of String)

        ''Private RenderingTier As Integer = RenderCapability.Tier >> 16
        Private WithEvents TrayNotify As New System.Windows.Forms.NotifyIcon()

        Private SearchText As TextBox
        Private WithEvents SearchPopup As Popup

        Private SabStartTimer As DispatcherTimer

        Private IcStopWait As Boolean = False
        Private WithEvents IC As ItemContainerGenerator

        Private PageSize As Integer = 250

        Private WithEvents SpotSource As SpotProvider
        Private WithEvents SpotOpener As BackgroundWorker

        Private WithEvents HeaderMenu As ContextMenu
        Private WithEvents SpotMenu As ContextMenu

        Private WithEvents SAB As New SabHelper
        Private WithEvents Suggest As New WebClient

        Public Sub New()

            Try

                InitializeComponent()

            Catch ex As Exception

                If Not (ex.InnerException Is Nothing) Then
                    Foutje(ex.InnerException.Message)
                Else
                    Foutje(ex.Message)
                End If

                Windows.Application.Current.Shutdown()

            End Try

        End Sub

        Private Sub Zoeken(ByVal bEnable As Boolean, ByVal bPic As Boolean)

            Dock.IsHitTestVisible = bEnable

            If bEnable Then

                EnableBijwerken()

            Else

                DisableBijwerken(bPic)

            End If

        End Sub

        Private Sub ClearFilter()

            If Not FilterSelected() Then Exit Sub

            LastSel = vbNullString

            For Each XL As FilterCat In FilterList.Items
                XL.Selected = False
            Next

            FilterList.SelectedIndex = -1

        End Sub

        Private Sub NoFilter(Optional ByVal bForce As Boolean = False, Optional ByVal WaitString As String = "Overzicht opbouwen...", Optional ByVal BlockUI As Boolean = True)

            ClearFilter()

            If (SpotSource.RowFilter <> sModule.DefaultFilter) Or (bForce) Then

                SetFilter(sModule.DefaultFilter, sModule.DefaultFilter, WaitString, True, True, BlockUI)

            End If

        End Sub

        Private Function GetFilterString(ByRef TheFilX As Collection) As String

            Dim SL As String

            If TheFilX Is Nothing Then Return vbNullString

            Dim QueryString As String = ""

            For Each SL In TheFilX  'Zoek Hoofdcats
                If SL.Substring(0, 1) = "X" Then
                    If Len(SL) = 3 Then
                        QueryString &= CStr(IIf(Len(QueryString) > 0, " OR ", "")) & Val(SL.Substring(1, 2)) + 1
                    End If
                End If
            Next

            QueryString &= CStr(IIf(Len(QueryString) > 0, " OR ", ""))

            For CC As Long = 0 To 9
                QueryString &= LimCat2(TheFilX, CC)
            Next

            If Microsoft.VisualBasic.Strings.Right(QueryString, 4) = " OR " Then
                QueryString = QueryString.Substring(0, QueryString.Length - 4)
            End If

            If Microsoft.VisualBasic.Strings.Right(QueryString, 5) = " AND " Then
                QueryString = QueryString.Substring(0, QueryString.Length - 5)
            End If

            Return RewriteQuery2("cats MATCH '" & QueryString.Trim & "'")

        End Function

        Private Function LimCat2(ByRef TheFilX As Collection, ByVal LimitCat As Long) As String

            Dim QueryString As String = ""

            Dim QueryA As String = LimCat(TheFilX, LimitCat, "a")
            If QueryA.Length > 0 Then QueryA += " AND "
            Dim QueryB As String = LimCat(TheFilX, LimitCat, "b")
            If QueryB.Length > 0 Then QueryB += " AND "
            Dim QueryC As String = LimCat(TheFilX, LimitCat, "c")
            If QueryC.Length > 0 Then QueryC += " AND "
            Dim QueryD As String = LimCat(TheFilX, LimitCat, "d")
            If QueryD.Length > 0 Then QueryD += " AND "
            Dim QueryZ As String = LimCat(TheFilX, LimitCat, "z")

            QueryString = QueryA & QueryB & QueryC & QueryD & QueryZ

            If Microsoft.VisualBasic.Strings.Right(QueryString, 5) = " AND " Then
                QueryString = QueryString.Substring(0, QueryString.Length - 5)
            End If

            If QueryString.Length > 0 Then

                If QueryString.Contains(" AND ") Then

                    QueryString = " (" & QueryString & ") OR "

                Else

                    QueryString = " " & QueryString & " OR "

                End If

            End If

            Return QueryString

        End Function

        Private Function LimCat(ByRef TheFilX As Collection, ByVal LimitCat As Long, ByVal LimitSub As String) As String

            Dim SL As String
            Dim xCount As Long

            If TheFilX Is Nothing Then Return vbNullString

            Dim QueryString As String = ""

            For Each SL In TheFilX ' Zoek Cat 
                If SL.Substring(0, 1) = "X" Then
                    If Len(SL) > 3 Then
                        If Val(SL.Substring(1, 2)) = LimitCat Then
                            If SL.Substring(3, 1).ToLower = LimitSub Then

                                xCount += 1
                                If xCount > 1 Then QueryString += " OR "

                                QueryString += (Val(SL.Substring(1, 2)) + 1) & LimitSub & Val(SL.Substring(4))

                            End If
                        End If
                    End If
                End If
            Next

            If xCount > 1 Then
                QueryString = "(" & QueryString & ")"
            End If

            Return QueryString

        End Function

        Private Function SupportsTaskProgress() As Boolean

            If System.Environment.OSVersion.Version.Major = 6 Then
                Return (System.Environment.OSVersion.Version.Minor >= 1)
            Else
                Return (System.Environment.OSVersion.Version.Major > 6)
            End If

        End Function

        Private Sub HasFilter(ByVal sName As String, Optional ByVal BlockUI As Boolean = True)

            If Not FilterSelected() Then Foutje("Filter Err")

            Dim TheFil As FilterCat

            LastSel = sName
            TheFil = GetFilter(sName)

            If Len(TheFil.Name) = 0 Or Len(TheFil.Query) = 0 Then Foutje("Filter Err2")

            If SpotSource.RowFilter <> TheFil.Query Then

                SetFilter(TheFil.Name, TheFil.Query, "Overzicht aan het filteren..", True, True, BlockUI)

                GebruikFilter.Content = SpotSource.QueryName

            End If

        End Sub

        'Private Sub ForceFilter(ByVal sName As String)

        '    If FilterList Is Nothing Then Exit Sub

        '    For Each XL As FilterCat In FilterList.Items
        '        If XL.Selected <> (XL.Name = sName) Then
        '            If (sName.Length > 0) Then
        '                If (XL.Name = sName) Then FilterList.SelectedItem = XL
        '            End If
        '        End If
        '    Next

        '    SwitchTab()

        '    If sName <> LastSel Then

        '        If Len(Trim$(sName)) > 0 Then

        '            HasFilter(sName)

        '        Else

        '            NoFilter()

        '        End If

        '    End If

        'End Sub

        Private Sub FilterList_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles FilterList.MouseUp

            Dim TempSel As String

            If FilterList Is Nothing Then Exit Sub
            If FilterList.SelectedItems Is Nothing Then Exit Sub

            If TypeOf e.OriginalSource Is ScrollViewer Then Exit Sub

            Dim dep As DependencyObject
            dep = CType(e.OriginalSource, DependencyObject)

            If TypeOf dep Is ScrollViewer Then Exit Sub

            If FilterList.SelectedItems.Count > 0 Then

                TempSel = ""

                For Each XL As FilterCat In FilterList.SelectedItems
                    TempSel = CStr(XL.Name)
                    Exit For
                Next

                For Each XL As FilterCat In FilterList.Items
                    If XL.Visible Then
                        If XL.Selected <> (XL.Name = TempSel) Then
                            If (TempSel.Length > 0) Then XL.Selected = (XL.Name = TempSel)
                        End If
                    End If
                Next

                If TempSel <> LastSel Then

                    If Len(Trim$(TempSel)) > 0 Then

                        HasFilter(TempSel)

                    Else

                        NoFilter()
                        Exit Sub

                    End If

                End If

            End If

            If e.ChangedButton = MouseButton.Right Then

                Try

                    Dim sName As String = ""

                    Dim X As ListBoxItem = CType(ItemsControl.ContainerFromElement(FilterList, dep), ListBoxItem)

                    If Not X Is Nothing Then
                        Dim sF As FilterCat = CType(X.Content, FilterCat)
                        If Not sF Is Nothing Then sName = sF.Name
                    End If

                    Dim ZX As ContextMenu = GetFilterMenu(sName, FilterList.SelectedIndex = 0, FilterList.SelectedIndex = FilterList.Items.Count - 1)
                    Dim fe As FrameworkElement = CType(e.Source, FrameworkElement)

                    fe.ContextMenu = ZX
                    fe.ContextMenu.IsOpen = True
                    fe.ContextMenu = Nothing

                Catch ex As Exception

                End Try

            End If

        End Sub

        Public Function GetFilterMenu(ByVal sFilter As String, ByVal bFirst As Boolean, ByVal bLast As Boolean) As ContextMenu

            Dim XM As New ContextMenu

            XM.FontFamily = Me.FontFamily
            XM.FontSize = Me.FontSize - 1
            XM.FontStyle = Me.FontStyle

            Dim mDelete As New MenuItem
            With mDelete
                .Header = "Verwijderen"
                .Tag = "DELETE " & sFilter
                .IsEnabled = Len(sFilter) > 0
                .Icon = GetIcon("delete")
                If Not .IsEnabled Then .Opacity = 0.5
                .AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf DoFilterMenu))
            End With

            Dim mEdit As New MenuItem
            With mEdit
                .Header = "Bewerken"
                .Tag = "EDIT " & sFilter
                .IsEnabled = False
                .Icon = GetIcon("settings")
                If Not .IsEnabled Then .Opacity = 0.5
                .AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf DoFilterMenu))
            End With

            Dim mS As New MenuItem
            With mS
                .Header = "Omhoog"
                .Tag = "UP " & sFilter
                .Icon = GetIcon("up")
                .IsEnabled = Not bFirst
                If Not .IsEnabled Then .Opacity = 0.5
                .AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf DoFilterMenu))
            End With

            Dim mS2 As New MenuItem
            With mS2
                .Header = "Omlaag"
                .Tag = "DOWN " & sFilter
                .Icon = GetIcon("down")
                .IsEnabled = Not bLast
                If Not .IsEnabled Then .Opacity = 0.5
                .AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf DoFilterMenu))
            End With

            Dim mReset As New MenuItem
            With mReset
                .Header = "Standaard"
                .Tag = "RESET"
                .Icon = GetIcon("gear")
                If Not .IsEnabled Then .Opacity = 0.5
                .AddHandler(MenuItem.ClickEvent, New RoutedEventHandler(AddressOf DoFilterMenu))
            End With

            XM.Items.Add(mS)
            XM.Items.Add(mS2)
            XM.Items.Add(New Separator)
            XM.Items.Add(mEdit)
            XM.Items.Add(mDelete)
            XM.Items.Add(New Separator)
            XM.Items.Add(mReset)

            Return XM

        End Function

        Private Sub DoFilterMenu(ByVal sender As Object, ByVal e As RoutedEventArgs)

            Dim zErr As String = ""

            Dim x As MenuItem = CType(e.OriginalSource, MenuItem)
            Dim x2 As ContextMenu = CType(x.Parent, ContextMenu)

            Select Case Split(CStr(x.Tag), " ")(0)
                Case "DELETE"

                    RemoveFilter(x.Tag.ToString.Substring(Len(Split(CStr(x.Tag), " ")(0)) + 1))

                Case "RESET"

                    If MsgBox("Weet je zeker dat je de standaard filters wilt terugzetten?", CType(MsgBoxStyle.YesNo + MsgBoxStyle.Information, MsgBoxStyle), "Filters") = MsgBoxResult.Yes Then
                        ReloadFilters(True)
                    End If

                Case "UP"

                    FilterDB.SwapFilter(x.Tag.ToString.Substring(Len(Split(CStr(x.Tag), " ")(0)) + 1), True)
                    ReloadFilters(False)

                Case "DOWN"

                    FilterDB.SwapFilter(x.Tag.ToString.Substring(Len(Split(CStr(x.Tag), " ")(0)) + 1), False)
                    ReloadFilters(False)

            End Select

        End Sub

        Private Sub RemoveFilter(ByVal sName As String)

            FilterDB.RemoveFilter(sName)

            Dim lp As Spotnet.FilterCat
            Dim bFound As Boolean

            For Each lp In FilterList.Items
                If lp.Visible Then
                    If Trim$(UCase$(lp.Name)) = Trim$(UCase$(sName)) Then
                        bFound = True
                        FilterList.Items.Remove(lp)
                        Exit For
                    End If
                End If
            Next

            If Not bFound Then Exit Sub
            If FilterList.Items.Count = 0 Then ReloadFilters(True) Else NoFilter()

        End Sub

        Private Sub ReloadFilters(ByVal bDefeault As Boolean)

            If bDefeault Then
                FilterDB.DefaultFilters("")
            End If

            FilterDB.LoadFilters()
            DisplayFilters()

            NoFilter()

        End Sub

        Private Function TreeChilds() As Boolean

            Dim fList As New List(Of FooViewModel)

            Dim LF As FooViewModel
            Dim LF2 As FooViewModel
            Dim LF3 As FooViewModel

            Dim SubCats As New Collection

            For lX = 0 To Me.tree.Items.Count - 1
                fList.Add(TryCast(Me.tree.Items(lX), FooViewModel))
            Next

            For Each LF In fList
                If (LF.IsChecked) Is Nothing Then
                    For Each LF2 In LF.Children
                        For Each LF3 In LF2.Children
                            If LF3.IsChecked Then
                                Return True
                            End If
                        Next
                    Next
                Else
                    If (LF.IsChecked) Then
                        Return True
                    End If
                End If
            Next

            Return False

        End Function

        Private Function NewFilter(ByVal sFiltername As String, ByVal zQuery As String, ByVal sImg As String, ByVal ShowErr As Boolean) As Boolean

            Dim sErr As String = ""

            If Len(Trim$(sFiltername)) = 0 Then Return False

            If FilterDB.AddFilter(sFiltername, zQuery, sImg, sErr) Then
                FilterDB.LoadFilters()
                FilterList.Items.Add(FilterDB.FilterOverview.Item(FilterDB.FilterOverview.Count - 1))
                Return True
            Else
                If ShowErr Then Foutje("Kan filter niet toevoegen: " & sErr)
                Return False
            End If

        End Function

        Friend Function SaveFilter(ByVal zQuery As String, ByVal sName As String, ByRef sError As String) As Boolean

            If FilterDB.FilterExist(sName) Then
                sError = ("Filter '" & sName & "' bestaat al!")
                Return False
            End If

            If Not NewFilter(sName, zQuery, "", False) Then
                sError = ("Kan filter '" & sName & "' niet toevoegen!")
                Return False
            End If

            Return True

        End Function

        Friend Sub SearchFilter(ByVal zQuery As String, ByVal sName As String)

            FirstTab(SpotSource.QueryName, "filter.ico")

            Me.Dispatcher.BeginInvoke(DispatcherPriority.Background, New SearchFilter3(AddressOf SearchFilter2), zQuery, sName)

        End Sub

        Delegate Sub SearchFilter3(ByVal zQuery As String, ByVal sName As String)
        Friend Sub SearchFilter2(ByVal zQuery As String, ByVal sName As String)

            SearchBox.Text = ""
            ClearFilter()
            SetFilter(SearchM & sName, zQuery, "Zoeken naar: " & sName & "..", True, True, True)

        End Sub

        Private Sub DisplayFilters()

            Dim TheFil As FilterCat

            If FilterList.Items.Count > 0 Then FilterList.Items.Clear()

            For Each TheFil In FilterDB.FilterOverview
                If GebruikFilter.Visibility = Visibility.Hidden Then
                    GebruikFilter.Content = TheFil.Name
                    GebruikFilter.Visibility = Visibility.Visible
                End If
                FilterList.Items.Add(TheFil)
            Next

        End Sub

        Friend Function GetFilter(ByVal sName As String) As FilterCat

            Dim TheFil As FilterCat

            For Each TheFil In FilterDB.FilterOverview
                If TheFil.Name = sName Then
                    Return TheFil
                End If
            Next

            Return Nothing

        End Function

        Private Sub ReopenTabs()

            Dim sMes As String = ""
            Dim sTitel As String = ""

            Try

                For Each XX As String In TabDB.Tabs
                    sMes = Split(XX, vbTab)(0)
                    sTitel = XX.Substring(sMes.Length + 1)
                    PrepareTab(sTitel, sMes)
                Next

            Catch ex As Exception

            End Try

        End Sub

        Friend Function DownloadsVisible() As Boolean

            Try
                If Me.WindowState = WindowState.Minimized Then Return False
                If Me.Visibility <> Visibility.Visible Then Return False
                If TabControl1.SelectedItem Is Nothing Then Return False
                Return IsDownload(CType(TabControl1.SelectedItem, TabItem))
            Catch
            End Try

            Return False

        End Function

        Friend Sub InitMain()

            Dim xErr As String = ""
            Dim HC As Boolean = False
            Dim PositionIndex As Long = 0

            Static DidOnce As Boolean = False

            Try

                If Not DidOnce Then

                    DidOnce = True

                    PositionIndex = 10

                    If SAB.External() Then
                        Map.IsEnabled = False
                        Map.Opacity = 0.5
                    End If

                    PositionIndex = 2
                    If My.Settings.GoogleSuggest Then HistoryDB.LoadHistory()

                    PositionIndex = 3
                    Dim iconStream As Stream = Application.GetResourceStream(New Uri("pack://application:,,,/Spotnet;component/Images/smallspotnet.ico")).Stream

                    PositionIndex = 4
                    TrayNotify.Icon = New System.Drawing.Icon(iconStream)
                    TrayNotify.Text = Me.Title

                    PositionIndex = 13

                    If (Not SpotSource.Connected) Or (Len(Trim$(ServersDB.oDown.Server)) = 0) Then

                        PositionIndex = 8

                        CS(250)
                        Mouse.OverrideCursor = Nothing

                        If Not SelecteerProvider() Then

                            Me.Close()
                            Exit Sub

                        End If

                        Mouse.OverrideCursor = Cursors.Wait

                        If Not OpenDB(xErr) Then

                            Foutje(xErr)
                            Me.Close()
                            Exit Sub

                        End If

                        NoFilter(True)
                        Me.UpdateLayout()

                    End If

                    PositionIndex = 44

                    If Not SpotSource.Connected Then

                        Foutje("Geen verbinding met database")

                        Me.Close()
                        Exit Sub

                    End If

                    UpdateStatus()

                    PositionIndex = 1
                    Downloads.ItemsSource = SAB.SabCol

                    PositionIndex = 11

                    If My.Settings.DownloadAction < 2 Then
                        If SAB.External() Or SAB.IsProcessRunning() Then
                            Dim zErr As String = ""
                            If Not SAB.StartSab(GetServer(ServerType.Download), True, zErr) Then
                                Foutje("Fout tijdens het starten van SABnzbd: " & zErr)
                            End If
                        Else
                            Dim zErr As String = ""
                            If Not SAB.LoadSab(GetServer(ServerType.Download), zErr) Then
                                Foutje("Fout tijdens het starten van SABnzbd: " & zErr)
                            End If
                        End If
                    End If

                    PositionIndex = 73

                    If My.Settings.AutoUpdate Or (My.Settings.DatabaseCount < 1) Then

                        CS(250)

                        If DoUpdate(xErr) Then Exit Sub

                        If (Len(xErr) > 0) And (xErr <> CancelMSG) Then
                            Foutje(xErr, "Fout")
                        End If

                    End If

                    EndWait(False)
                    CS()

                End If

            Catch ex As Exception

                Foutje("Load: Code #" & PositionIndex & ": " & ex.Message)
                Me.Close()

            End Try

        End Sub

        Private Sub MainWindow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing

            Try

                Try
                    If Me.Visibility = Visibility.Visible Then

                        My.Settings.WindowTop = Me.Top
                        My.Settings.WindowLeft = Me.Left
                        My.Settings.WindowWidth = Me.Width
                        My.Settings.WindowHeight = Me.Height
                        My.Settings.WindowMaximized = (Me.WindowState = WindowState.Maximized)
                        My.Settings.Save()

                    End If
                Catch
                End Try

                CancelSuggest()
                TrayNotify = Nothing
                ClearPhuses()

                Me.Visibility = Windows.Visibility.Hidden

                StopDownloads()
                SAB = Nothing

            Catch ex As Exception
                Foutje("Closed: " & ex.Message)

            End Try

        End Sub

        Private Sub StopDownloads()

            Dim zErr As String = ""

            If Not SAB Is Nothing Then
                If Not SAB.CloseSab(zErr) Then Foutje("Fout tijdens het afsluiten van SABnzbd: " & zErr)
            End If

        End Sub

        Private Function OpenDB(ByRef zErr As String) As Boolean

            Try

                Dim sName As String = SpotSource.QueryName
                Dim sFilter As String = SpotSource.RowFilter

                SpotSource.Close()

                If Not SpotSource.Connect(GetDBFilename("dbs"), zErr) Then
                    Return False
                End If

                SpotSource.QueryName = sName
                SpotSource.RowFilter = sFilter

                Return True

            Catch ex As Exception

                zErr = "OpenDB: " & ex.Message
                Return False

            End Try

        End Function

        Private Sub MainWindow_Initialized(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Initialized

            Dim lPosIndex As Long = 1

            Try

                lPosIndex = 1
                Mouse.OverrideCursor = Cursors.Wait

                lPosIndex = 22

                My.Settings.SortColumn = "rowid"
                My.Settings.SortDirection = "DESC"

                Dim ZK() As Byte = MakeLatin(My.Settings.Columns)
                Dim OneVisible As Boolean = False

                For ZL As Integer = 0 To UBound(ZK)
                    If ZL < Spots.Columns.Count Then
                        If ZK(ZL) > 48 Then
                            OneVisible = True
                            Exit For
                        End If
                    End If
                Next

                If Not OneVisible Then
                    ZK = MakeLatin("1234005000")
                End If

                For ZL As Integer = 0 To UBound(ZK)
                    If ZL < Spots.Columns.Count Then
                        If ZK(ZL) > 48 Then
                            Spots.Columns(ZL).Visibility = Visibility.Visible
                            Spots.Columns(ZL).DisplayIndex = Val(Chr(ZK(ZL))) - 1
                        Else
                            Spots.Columns(ZL).Visibility = Visibility.Hidden
                        End If
                    End If
                Next

                UpdateSortCol()

                lPosIndex = 3

                If My.Settings.WindowWidth <> 0 Then

                    Me.Width = My.Settings.WindowWidth
                    Me.Height = My.Settings.WindowHeight
                    Me.Top = My.Settings.WindowTop
                    Me.Left = My.Settings.WindowLeft

                    Me.WindowStartupLocation = WindowStartupLocation.Manual

                    If Me.Top < 0 Then Me.Top = 0
                    If Me.Left < 0 Then Me.Left = 0

                    If Me.Top + Me.Height > (My.Computer.Screen.Bounds.Bottom + 10) Then
                        Me.Top = 0
                        Me.Height = My.Computer.Screen.Bounds.Bottom - 50
                        Me.WindowStartupLocation = WindowStartupLocation.CenterScreen
                    End If

                    If Me.Left + Me.Width > (My.Computer.Screen.Bounds.Right + 10) Then
                        Me.Left = 0
                        Me.Width = My.Computer.Screen.Bounds.Right - 50
                        Me.WindowStartupLocation = WindowStartupLocation.CenterScreen
                    End If

                Else

                    If My.Computer.Screen.Bounds.Bottom * 0.9 > 600 Then
                        Me.Height = My.Computer.Screen.Bounds.Bottom * 0.9
                    Else
                        Me.Width = My.Computer.Screen.Bounds.Bottom - 50
                        My.Settings.WindowMaximized = True

                    End If

                    If My.Computer.Screen.Bounds.Right * 0.9 > 800 Then
                        Me.Width = My.Computer.Screen.Bounds.Right * 0.9
                    Else
                        Me.Width = My.Computer.Screen.Bounds.Right - 50
                        My.Settings.WindowMaximized = True
                    End If

                    Me.WindowStartupLocation = WindowStartupLocation.CenterScreen

                End If

                If My.Settings.WindowMaximized Then
                    Me.WindowState = WindowState.Maximized
                Else
                    Me.WindowState = WindowState.Normal
                End If

                lPosIndex = 4

                If Not ServersDB.LoadServers() Then
                    Foutje("Kan servers.xml niet laden!")
                    Me.Close()
                    Exit Sub
                End If

                lPosIndex = 5

                If Not FilterDB.LoadFilters() Then
                    Foutje("Kan filters.xml niet laden!")
                    Me.Close()
                    Exit Sub
                End If

                lPosIndex = 6
                GebruikFilter.Visibility = Visibility.Hidden
                DisplayFilters()

                lPosIndex = 10

                UpdateFonts()
                ShowDownloads(My.Settings.DownloadAction < 2, False)
                Uitgebreid.IsChecked = My.Settings.AdvancedSearch

                lPosIndex = 11

                If My.Settings.SaveTabs Then
                    TabDB.LoadTabs()
                    ReopenTabs()
                End If

                lPosIndex = 12
                WebRequest.DefaultWebProxy = Nothing
                CheckSpotMenu()

                lPosIndex = 13

                Dim sErr As String = ""

                SpotSource = New SpotProvider()

                If (Len(Trim$(ServersDB.oDown.Server)) > 0) Then

                    If Not SpotSource.Connect(GetDBFilename("dbs"), sErr) Then

                        If SpotSource.Corrupted Then
                            Try
                                File.Delete(GetDBFilename("dbs"))
                            Catch
                            End Try
                        End If

                        Foutje(sErr)
                        Me.Close()

                        Exit Sub

                    End If

                End If

                Zoeken(True, False)

                SpotSource.QueryName = sModule.DefaultFilter
                SpotSource.RowFilter = sModule.DefaultFilter

                lPosIndex = 14

                IC = Spots.ItemContainerGenerator
                Spots.ItemsSource = New DataVirtualization.VirtualList(Of SpotRow)(SpotSource, PageSize, SynchronizationContext.Current)

                lPosIndex = 16

                UpdateStatus()
                Dock.IsHitTestVisible = False

                Exit Sub

            Catch ex2 As NullReferenceException

                If lPosIndex = 14 Then

                    If Not SpotSource Is Nothing Then
                        If SpotSource.Corrupted Then
                            Try
                                File.Delete(GetDBFilename("dbs"))
                            Catch
                            End Try
                        End If
                    End If

                    Me.Close()
                    Exit Sub

                End If

                Foutje("Initialized: Code# " & lPosIndex & ": " & ex2.Message)

            Catch ex As Exception

                Foutje("Initialized: Code# " & lPosIndex & ": " & ex.Message)

            End Try

            Me.Close()

        End Sub

        Private Sub UpdateSortCol()

            Dim bFound As Boolean = False

            For ZL As Integer = 0 To Spots.Columns.Count - 1
                If CStr(Spots.Columns(ZL).Header).ToLower = TranslateCol2(My.Settings.SortColumn) Then
                    bFound = True
                    If My.Settings.SortDirection.ToUpper = "ASC" Then
                        Spots.Columns(ZL).SortDirection = ComponentModel.ListSortDirection.Ascending
                    Else
                        Spots.Columns(ZL).SortDirection = ComponentModel.ListSortDirection.Descending
                    End If
                    Exit For
                End If
            Next

            If Not bFound Then Foutje("Column " & My.Settings.SortColumn & " not found?")

        End Sub

        Friend Function GetDBFilename(ByVal sExtension As String) As String

            Try

                Dim sFile As String = SafeName(Trim$(ServersDB.oDown.Server)).ToLower

                If Len(sFile) = 0 Then Return vbNullString

                Return SettingsFolder() & "\" & sFile & "." & sExtension

            Catch ex As Exception
            End Try

            Return vbNullString

        End Function

        Private Sub tree_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles tree.MouseMove

            btnToevoegen.IsEnabled = TreeChilds() Or (Len(Trim(Filterbox.Text)) > 0)
            CheckToevoegen()

        End Sub

        Private Sub tree_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles tree.MouseUp

            btnToevoegen.IsEnabled = TreeChilds() Or (Len(Trim(Filterbox.Text)) > 0)
            CheckToevoegen()

        End Sub

        Private Sub tree_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.Windows.RoutedPropertyChangedEventArgs(Of System.Object)) Handles tree.SelectedItemChanged

            btnToevoegen.IsEnabled = TreeChilds() Or (Len(Trim(Filterbox.Text)) > 0)
            CheckToevoegen()

        End Sub

        Private Sub MainToolBar_Initialized(ByVal sender As Object, ByVal e As System.EventArgs) Handles MainToolBar.Initialized

            Dim a As FrameworkElement

            ToolBarTray.SetIsLocked(MainToolBar, True)

            For Each a In MainToolBar.Items
                ToolBar.SetOverflowMode(a, OverflowMode.Never)
            Next

        End Sub

        Private Function UpdateDB(ByRef sErr As String, ByRef rReturn As rSaveSpots) As Boolean

            Dim StatusEvents As Status
            Dim rSpot As rSaveSpots = Nothing
            Dim rComment As rSaveComments = Nothing

            Try

                StatusEvents = New Status

                AddHandler StatusEvents.StatusChanged, AddressOf StatusChanged

                StatusEvents.HeadersFile = GetDBFilename("dbs")
                StatusEvents.CommentsFile = GetDBFilename("dbc")

                StatusEvents.Owner = Me
                StatusEvents.Title = "Database bijwerken..."

                StatusEvents.ShowDialog()

                rSpot = StatusEvents.HeaderResults
                rComment = StatusEvents.CommentResults

                If StatusEvents.HasError Then
                    If (rSpot Is Nothing) And (rComment Is Nothing) Then

                        sErr = StatusEvents.LastError
                        StatusEvents = Nothing

                        Return False

                    Else

                        Me.Cursor = Nothing
                        Mouse.OverrideCursor = Nothing
                        Foutje(StatusEvents.LastError, "Fout")

                    End If
                End If

                StatusEvents = Nothing

                rReturn = rSpot
                Return True

            Catch ex As Exception

                sErr = "UpdateDB: " & ex.Message
                Return False

            End Try

        End Function

        Private Sub DoBlur()

            'Dim bGreyed As Boolean = True

            'If RenderingTier > 1 Or System.Diagnostics.Debugger.IsAttached Then
            '    If Me.Effect Is Nothing Then

            '        If (TabControl1.SelectedIndex = 0) Or IsDownload(TabControl1.SelectedItem) Then

            '            Dim blur As New BlurEffect
            '            blur.Radius = 4

            '            If bGreyed Then

            '                Dim mask As SolidColorBrush = New SolidColorBrush(Colors.White)
            '                mask.Opacity = 0.8
            '                Me.OpacityMask = mask

            '            End If

            '            Me.Effect = blur

            '        End If

            '    End If
            'End If

        End Sub

        Private Sub ClearBlur()

            'Me.Effect = Nothing
            'Me.OpacityMask = Nothing

        End Sub

        Private Sub Spots_ContextMenuOpening(ByVal sender As Object, ByVal e As System.Windows.Controls.ContextMenuEventArgs) Handles Spots.ContextMenuOpening

            Dim bFound As Boolean = False

            e.Handled = True
            e.OriginalSource.ContextMenu = Nothing

            Spots.ContextMenu = Nothing

            If Not TypeOf e.OriginalSource Is ScrollViewer Then
                If Not TypeOf e.OriginalSource Is Microsoft.Windows.Themes.DataGridHeaderBorder Then
                    If TypeOf e.OriginalSource Is TextBlock Then
                        Dim Lx As TextBlock = CType(e.OriginalSource, TextBlock)
                        For Each Xz As DataGridColumn In Spots.Columns
                            If CStr(Xz.Header).ToLower = Lx.DataContext.ToString.ToLower Then
                                bFound = True
                                Exit For
                            End If
                        Next
                    End If
                Else
                    bFound = True
                End If
            Else
                bFound = (e.CursorTop < 50)
            End If

            If bFound Then
                LoadHeaderMenu()
                e.OriginalSource.ContextMenu = HeaderMenu
                e.OriginalSource.ContextMenu.IsOpen = True
                e.OriginalSource.ContextMenu = Nothing
                Exit Sub
            End If

            Dim bDoContext As Boolean
            Dim bDoExtContext As Boolean

            If ((TypeOf e.OriginalSource Is TextBlock)) Then
                Dim zx As TextBlock = CType(e.OriginalSource, TextBlock)
                If TypeOf zx.Parent Is DataGridCell Then
                    bDoContext = True
                    bDoExtContext = True
                End If
            End If

            If ((TypeOf e.OriginalSource Is Border)) Then
                Dim zx As Border = CType(e.OriginalSource, Border)
                If TypeOf zx.Parent Is DataGridCell Then
                    bDoContext = True
                    bDoExtContext = True
                End If
            End If

            If TypeOf e.OriginalSource Is ScrollViewer Then
                bDoContext = IsSearching()
            End If

            If bDoContext Then

                SpotMenu = New ContextMenu

                SpotMenu.FontFamily = Me.FontFamily
                SpotMenu.FontSize = Me.FontSize - 1
                SpotMenu.FontStyle = Me.FontStyle

                Dim tk As New MenuItem

                If Not SelectedSpot() Is Nothing Then
                    If bDoExtContext Then

                        tk = New MenuItem

						tk.Tag = "report"
						tk.Header = "Melding versturen"
						tk.Icon = GetIcon("warning")

                        SpotMenu.Items.Add(tk)
                        SpotMenu.Items.Add(New Separator)

                        tk = New MenuItem

                        tk.Tag = "fav"

                        tk.IsEnabled = (Len(SelectedSpot.Modulus) > 0) And (Not BlackList.Contains(SelectedSpot.Modulus))

                        If WhiteList.Contains(SelectedSpot.Modulus) Then
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
                        tk.IsEnabled = (Len(SelectedSpot.Modulus) > 0) And (Not WhiteList.Contains(SelectedSpot.Modulus)) And (SelectedSpot.Modulus <> GetModulus())

                        If BlackList.Contains(SelectedSpot.Modulus) Then
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
                        SpotMenu.Items.Add(New Separator)

                    End If
                End If

                tk = New MenuItem
                tk.Header = "Zoekopdracht opslaan"
                tk.Icon = GetIcon("save")
                tk.IsEnabled = IsSearching()

                If tk.IsEnabled Then
                    tk.Icon.Opacity = 1
                Else
                    tk.Icon.Opacity = 0.5
                End If

                tk.Tag = "filter"

                SpotMenu.Items.Add(tk)

                Spots.ContextMenu = SpotMenu
                Spots.ContextMenu.IsOpen = True

            End If

            Spots.ContextMenu = Nothing

        End Sub

        Private Sub Spots_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Spots.PreviewKeyDown

            If e.Key = Key.Enter Then
                e.Handled = True
                OpenSel()
                Exit Sub
            End If

            If e.Key = Key.Delete Then
                e.Handled = True
                Exit Sub
            End If

            e.Handled = False

        End Sub

        Private Sub OpenAbout()

            If Not TabControl1.SelectedItem Is Nothing Then
                If TypeOf TabControl1.SelectedItem.Content Is AboutControl Then Exit Sub
            End If

            Dim newTab As New CloseableTabItem

            newTab.Header = CreateHeader("Over", "about.ico")

            Dim KL As New AboutControl
            newTab.Content = KL

            TabControl1.Items.Add(newTab)

        End Sub

        Private Sub OpenAdd()

            If Not TabControl1.SelectedItem Is Nothing Then
                If TypeOf TabControl1.SelectedItem.Content Is Toevoegen Then Exit Sub
            End If

            Dim newTab As New CloseableTabItem

            newTab.Header = CreateHeader("Toevoegen", "addspot.ico")

            Dim KL As New Toevoegen

            newTab.Content = KL

            TabControl1.Items.Add(newTab)

        End Sub

        Public Function DownloadNZB(ByVal sLoc As String, ByVal sTitle As String) As Boolean

            Dim zErr As String = ""

            DoWait("NZB downloaden...")

            If Not GUIStartSab() Then
                EndWait(True)
                Return False
            End If

            If Not SAB.AddNZB(sLoc, sTitle, zErr) Then
                Foutje(zErr)
                EndWait(True)
                Return False
            End If

            EndWait(True)
            Return True

        End Function

        Friend Sub OpenURL(ByVal sUrl As String, ByVal sTitle As String, ByVal ExTab As CloseableTabItem)

            Dim newTab As CloseableTabItem

            If Not ExTab Is Nothing Then
                newTab = ExTab
            Else
                newTab = New CloseableTabItem
            End If

            If ExTab Is Nothing Then newTab.Header = CreateHeader(sTitle, "url.ico")

            Dim zx As New HTMLView

            newTab.Content = zx

            Dim tW As System.Windows.Forms.WebBrowser

            tW = zx.GetWb(True, "")

            With tW

                .Navigate(AddHttp(sUrl), False)

            End With

            Dim Kl As New UrlInfo
            Kl.Title = sTitle
            Kl.URL = AddHttp(sUrl)
            Kl.TabLoaded = True

            newTab.Tag = Kl

            If ExTab Is Nothing Then

                If My.Settings.SaveTabs Then SaveTabs(sUrl.Replace(vbTab, "") & vbTab & sTitle.Replace(vbTab, ""))

                TabControl1.Items.Add(newTab)

            End If

        End Sub

        Private Sub PrepareTab(ByVal sTitle As String, ByVal sLoc As String)

            Try

                Dim newTab As New CloseableTabItem

                newTab.AutoSelect = False
                newTab.Header = CreateHeader(System.Net.WebUtility.HtmlDecode(sTitle), "smallspotnet.ico")

                If sLoc.StartsWith("<") Then

                    Dim DK As New SpotInfo

                    DK.Spot = New SpotEx
                    DK.Spot.Title = sTitle
                    DK.Spot.MessageID = MakeMsg(sLoc, False)
                    DK.TabLoaded = False

                    newTab.Tag = DK

                Else

                    Dim DU As New UrlInfo

                    DU.Title = sTitle
                    DU.TabLoaded = False
                    DU.URL = sLoc

                    newTab.Tag = DU

                End If

                TabControl1.Items.Add(newTab)

            Catch ex As Exception
                Foutje("PrepareTab: " & ex.Message)
            End Try

        End Sub

        Private Sub OpenSpot(ByVal xArticle As Long, ByVal xMsgID As String, ByRef ExTab As CloseableTabItem)

            Try

                Dim zErr As String = ""

                If Len(xMsgID) = 0 Then Exit Sub

                xMsgID = MakeMsg(xMsgID)

                If ExTab Is Nothing Then
                    If Not GetTab(xMsgID) Is Nothing Then
                        Dim Zl As CloseableTabItem = CType(GetTab(xMsgID), CloseableTabItem)
                        Zl.IsSelected = True
                        Exit Sub
                    End If
                End If

                DoWait("Spot aan het openen..")

                SpotOpener = New BackgroundWorker
                SpotOpener.WorkerSupportsCancellation = False

                Dim xObj(2) As Object

                xObj(0) = xMsgID
                xObj(1) = ExTab
                xObj(2) = xArticle

                SpotOpener.RunWorkerAsync(xObj)

                Exit Sub

            Catch ex As Exception
                Foutje("OpenSpot: " & ex.Message)
            End Try

        End Sub

        Private Sub OpenSpot2(ByVal xSpot As SpotEx, ByRef ExTab As CloseableTabItem)

            Dim zx As New HTMLView
            Dim newTab As CloseableTabItem
            Dim MaxSize As Integer = 350
            Dim tW As System.Windows.Forms.WebBrowser

            Try

                If BlackList.Contains(xSpot.User.Modulus) Then
                    Throw New Exception("Afzender staat op de blacklist!")
                End If

                If Not ExTab Is Nothing Then
                    newTab = ExTab
                    zx.Visibility = Visibility.Hidden
                Else
                    newTab = New CloseableTabItem
                    newTab.Header = CreateHeader(System.Net.WebUtility.HtmlDecode(xSpot.Title), "smallspotnet.ico")
                End If

                newTab.Content = zx

                tW = zx.GetWb(False, GetDBFilename("dbc"))

                If My.Computer.Keyboard.ShiftKeyDown Then
                    xSpot.Image = ""
                    xSpot.ImageID = ""
                End If

                If xSpot.ImageWidth > 0 Or xSpot.ImageHeight > 0 Then
                    If xSpot.ImageWidth > MaxSize Then xSpot.ImageWidth = MaxSize
                    If xSpot.ImageHeight < 32 Then xSpot.ImageWidth = 32
                    If xSpot.ImageHeight > MaxSize Then xSpot.ImageHeight = MaxSize
                    If xSpot.ImageHeight < 32 Then xSpot.ImageHeight = 32
                End If

                Dim DK As New SpotInfo

                DK.Spot = xSpot
                DK.TabLoaded = True

                newTab.Tag = DK

                tW.DocumentText = SpotParser.ParseSpot(xSpot, CByte(My.Settings.FontSize + 2))

            Catch ex As Exception

                Foutje(ex.Message, "Fout")

                If Not ExTab Is Nothing Then
                    TabControl1.Items.Remove(ExTab)
                End If

                EndWait(True)
                Exit Sub

            End Try

            Try

                If ExTab Is Nothing Then
                    If My.Settings.SaveTabs Then SaveTabs(MakeMsg(xSpot.MessageID) & vbTab & xSpot.Title.Replace(vbTab, ""))
                    TabControl1.Items.Add(newTab)
                Else
                    zx.Visibility = Visibility.Visible
                End If

            Catch ex As Exception

                Foutje("OpenSpot: " & ex.Message)

            End Try

            EndWait(True)

        End Sub

        Private Sub AddReport(ByVal sMsgID As String, ByVal sTitle As String)

            Dim zErr As String = ""

            If Len(sMsgID) = 0 Then Exit Sub

            Dim zDesc As String = InputBox("Waarom moet deze spot verwijderd worden?", "Spot melden")

            If Len(zDesc.Trim) = 0 Then
                Exit Sub
            End If

            DoWait("Melding plaatsen...")

            Try

                If SpotClient.Spots.CreatReport(UploadPhuse, StripNonAlphaNumericCharacters(My.Settings.Nickname), zDesc, My.Settings.ReportGroup, sMsgID, sTitle, zErr) Then

                    EndWait(True)
                    MsgBox("Je melding is verzonden, bedankt voor de moeite.", MsgBoxStyle.Information, "Verzonden")
                    Exit Sub

                End If

            Catch ex As Exception
                zErr = ex.Message
            End Try

            Foutje(zErr)
            EndWait(True)

        End Sub

        Private Sub DeleteArticle(ByVal sMsgID As String, ByVal sTitle As String)

            Dim zErr As String = ""

            If Len(sMsgID) = 0 Then Exit Sub

            DoWait("Spot verwijderen...")

            Try

                If SpotClient.Spots.DeleteArticle(UploadPhuse, My.Settings.Nickname, sTitle, My.Settings.HeaderGroup, sMsgID, "", zErr) Then

                    EndWait(True)
                    Exit Sub

                End If

            Catch ex As Exception
                zErr = ex.Message
            End Try

            Foutje(zErr)
            EndWait(True)

        End Sub

        Private Sub OpenSel()

            If Spots.SelectedItems.Count = 1 Then

                Dim Zk As SpotRow = SelectedSpot()

                Try

                    OpenSpot(Zk.ID, MakeMsg(SpotSource.GetMessageID("spots", Zk.ID)), Nothing)

                Catch ex As Exception

                    Foutje("OpenTab: " & ex.Message)

                End Try

            End If

        End Sub

        Private Function SelectedSpot() As SpotRow

            If Spots.SelectedItems.Count = 1 Then

                Try

                    Return CType(Spots.SelectedItem, VirtualListItem(Of SpotRow)).Data

                Catch ex As Exception
                End Try

            End If

            Return Nothing

        End Function

        Private Function GetTab(ByVal sMes As String) As TabItem

            Dim DK As SpotInfo

            For Each zx2 As TabItem In TabControl1.Items
                If TypeOf zx2.Tag Is SpotInfo Then
                    Try
                        DK = CType(zx2.Tag, SpotInfo)
                        If MakeMsg(DK.Spot.MessageID).Trim.ToLower = MakeMsg(sMes).Trim.ToLower Then
                            Return zx2
                        End If
                    Catch
                    End Try
                End If
            Next

            Return Nothing

        End Function

        Public Sub SaveTabs(Optional ByVal zExtra As String = "")

            Dim ZK As UrlInfo
            Dim DK As SpotInfo

            Dim THeTabs As New List(Of String)

            For Each zx2 As TabItem In TabControl1.Items
                If TypeOf zx2.Tag Is SpotInfo Then
                    Try
                        DK = CType(zx2.Tag, SpotInfo)
                        If Len(DK.Spot.MessageID) > 0 Then
                            THeTabs.Add(MakeMsg(DK.Spot.MessageID) & vbTab & DK.Spot.Title.Replace(vbTab, ""))
                        End If
                    Catch

                    End Try
                End If
                If TypeOf zx2.Tag Is UrlInfo Then
                    Try
                        ZK = CType(zx2.Tag, UrlInfo)
                        If Len(ZK.URL) > 0 Then
                            THeTabs.Add(ZK.URL.Replace(vbTab, "") & vbTab & ZK.Title.Replace(vbTab, ""))
                        End If
                    Catch ex As Exception

                    End Try
                End If
            Next

            If Len(zExtra) > 0 Then THeTabs.Add(zExtra)
            TabDB.SaveTabs(THeTabs)

        End Sub

        Private Sub UpdateStatus()

            Try

                If SpotSource Is Nothing Then
                    Throw New Exception("SpotSource Is Nothing")
                End If

                Dim TotalSpots As Long = My.Settings.DatabaseCount

                If (SpotSource.QueryCount < 1) And (TotalSpots > 0) Then

                    StatusChanged("Geen spots gevonden", -1)

                Else

                    If (SpotSource.QueryCount = TotalSpots) Or (SpotSource.RowFilter = sModule.DefaultFilter) Or (Len(SpotSource.RowFilter) = 0) Then

                        Select Case TotalSpots
                            Case Is < 1
                                StatusChanged("Geen spots in de database", -1)

                            Case 1
                                StatusChanged("1 spot in de database", -1)

                            Case Else
                                StatusChanged(Replace$(TotalSpots.ToString("#,#"), ",", ".") & " spots in de database", -1)

                        End Select

                    Else

                        Select Case SpotSource.QueryCount
                            Case Is < 1
                                StatusChanged("Geen spots gevonden", -1)
                            Case 1
                                StatusChanged("1 spot (van de " & TotalSpots.ToString("#,#").Replace(",", ".") & ")", -1)
                            Case Else
                                StatusChanged(Replace$(SpotSource.QueryCount.ToString("#,#"), ",", ".") & " spots (van de " & TotalSpots.ToString("#,#").Replace(",", ".") & ")", -1)
                        End Select

                    End If

                End If

            Catch ex As Exception
                Foutje("UpdateStatus: " & ex.Message)
            End Try

        End Sub

        Private Sub ChangeName(ByVal sHeader As String, ByVal sIcon As String)

            Dim dk As TabItem
            dk = CType(TabControl1.Items(0), TabItem)
            If dk Is Nothing Then Exit Sub

            If GetHeader(dk.Header).Trim <> sHeader.Trim Then
                If Len(sIcon) > 0 Then
                    dk.Header = CreateHeader(sHeader, sIcon, Nothing)
                Else
                    dk.Header = CreateHeader(sHeader, sIcon, GetHeaderIcon(dk.Header))
                End If
            End If

        End Sub

        Private Sub FirstTab(ByVal sHeader As String, ByVal sIcon As String)

            Dim dk As TabItem
            dk = CType(TabControl1.Items(0), TabItem)
            If dk Is Nothing Then Exit Sub

            If Not TabControl1.SelectedItem Is dk Then
                ChangeName(sHeader, sIcon)
                TabControl1.SelectedItem = dk
            End If

        End Sub

        Private Sub ShowDownloads(ByVal bVisible As Boolean, ByVal bStartStop As Boolean)

            Dim dk As TabItem
            dk = CType(TabControl1.Items(1), TabItem)

            If bVisible Then
                dk.Visibility = Windows.Visibility.Visible
            Else
                dk.Visibility = Windows.Visibility.Collapsed
            End If

            If bStartStop Then
                If Not bVisible Then
                    If IsDownload(CType(TabControl1.SelectedItem, TabItem)) Then
                        TabControl1.SelectedIndex = 0
                    End If
                    StopDownloads()
                End If
            End If

        End Sub

        Friend Sub DoWait(ByVal sMsg As String, Optional ByVal ZoekPlaatje As Boolean = False)

            Me.Cursor = System.Windows.Input.Cursors.Wait

            Mouse.OverrideCursor = Cursors.Wait

            Zoeken(False, ZoekPlaatje)

            StatusChanged(sMsg, 0)

        End Sub

        Friend Sub EndWait(ByVal bUpdateStatus As Boolean)

            Zoeken(True, False)

            Mouse.OverrideCursor = Nothing
            Me.Cursor = Nothing

            If bUpdateStatus Then UpdateStatus()

        End Sub

        Private Function DoUpdate(ByRef xErr As String) As Boolean

            Dim OldNew As Long = -1
            Dim EmptyCats(10) As Integer
            Dim bRes As rSaveSpots = Nothing
            Dim MustResetNew As Boolean = False
            Dim SetNew As Long = My.Settings.DatabaseMax

            Try

                DisableBijwerken(False)
                DoWait("Database bijwerken...", True)

                OldNew = SpotSource.RowNew
                MustResetNew = OldNew > 1

                If Not UpdateDB(xErr, bRes) Then

                    If Not SpotSource.Connected Then
                        OpenDB("")
                        SpotSource.RowNew = OldNew
                    End If

                    EnableBijwerken()
                    EndWait(True)

                    Return False

                End If

                If Not SpotSource.Connected Then

                    StatusChanged("Database openen...", 0)

                    If Not OpenDB(xErr) Then

                        EnableBijwerken()
                        EndWait(True)

                        Return False

                    End If

                End If

                Dim SpotsAdded As Boolean = False
                Dim SpotsRemoved As Boolean = False

                If Not bRes Is Nothing Then

                    SpotsAdded = (bRes.SpotsAdded > 0)
                    SpotsRemoved = (bRes.SpotsDeleted > 0)

                End If

                If SpotsAdded Then

                    StatusChanged("Overzicht opbouwen...", 0)

                    SpotSource.RowNew = SetNew

                    NoFilter(True, , False)

                    UpdateNewCats(bRes.NewCats)
                    StatusChanged(NewMsg(bRes.SpotsAdded), -1)

                    StopWait(False)

                    Return True

                Else

                    If MustResetNew Or SpotsRemoved Then

                        StatusChanged("Overzicht opbouwen...", 0)

                        If MustResetNew Then SpotSource.RowNew = -1

                        ReloadFilter(SpotsRemoved, False, , False)

                        If MustResetNew Then UpdateNewCats(EmptyCats)

                    End If

                    StatusChanged(NewMsg(0), -1)

                    EnableBijwerken()
                    EndWait(False)

                    Return True

                End If

            Catch ex As Exception

                xErr = "DoUpdate: " & ex.Message

                EnableBijwerken()
                EndWait(True)

                Return False

            End Try

        End Function

        Private Sub DisableBijwerken(ByVal bPic As Boolean)

            If bPic Then ZoekPlaatje.Opacity = 0.3
            ZoekPlaatje.ToolTip = "Spots worden geladen..."
            ZoekPlaatje.Cursor = Nothing

            Provider.IsEnabled = False
            Provider.Opacity = 0.5

            Bijwerken.IsEnabled = False
            Bijwerken.Opacity = 0.5

            Bij.Source = GetImage("refresh2.png")
            Bij.Opacity = 0.8
            Bij.IsEnabled = False
            Bij.ToolTip = ""
            Bij.Cursor = Nothing

        End Sub

        Private Sub EnableBijwerken()

            ZoekPlaatje.Opacity = 1
            ZoekPlaatje.ToolTip = "Zoeken..."
            ZoekPlaatje.Cursor = Windows.Input.Cursors.Hand

            Provider.IsEnabled = True
            Provider.Opacity = 1

            Bijwerken.IsEnabled = True
            Bijwerken.Opacity = 1

            Bij.Source = GetImage("refresh.png")
            Bij.Opacity = 1
            Bij.IsEnabled = True
            Bij.ToolTip = "Bijwerken"
            Bij.Cursor = Cursors.Hand

        End Sub

        Private Function RewriteQuery(ByVal sIn As String) As String

            If Len(sIn) = 0 Then Return sIn

            Select Case sIn.Replace(" ", "").Replace("(", "").Replace(")", "").ToLower
                Case "cat=1", "searchmatch'cats:1'"
                    Return "cats MATCH '1'"
                Case "cat=2", "searchmatch'cats:2'"
                    Return "cats MATCH '2'"
                Case "cat=3", "searchmatch'cats:3'"
                    Return "cats MATCH '3'"
                Case "cat=4", "searchmatch'cats:4'"
                    Return "cats MATCH '4'"
                Case "cat=5", "searchmatch'cats:5'"
                    Return "cats MATCH '5'"
                Case "cat=6", "searchmatch'cats:6'"
                    Return "cats MATCH '6'"
                Case "cat=9", "searchmatch'cats:9'"
                    Return "cats MATCH '9'"
            End Select

            Return sIn

        End Function

        Private Function RewriteQuery2(ByVal sIn As String) As String

            If Len(sIn) = 0 Then Return sIn

            Select Case sIn.Replace(" ", "").Replace("(", "").Replace(")", "").ToLower
                Case "searchmatch'cats:1'", "catsmatch'1'"
                    Return "cat = 1"
                Case "searchmatch'cats:2'", "catsmatch'2'"
                    Return "cat = 2"
                Case "searchmatch'cats:3'", "catsmatch'3'"
                    Return "cat = 3"
                Case "searchmatch'cats:4'", "catsmatch'4'"
                    Return "cat = 4"
                Case "searchmatch'cats:5'", "catsmatch'5'"
                    Return "cat = 5"
                Case "searchmatch'cats:6'", "catsmatch'6'"
                    Return "cat = 6"
                Case "searchmatch'cats:9'", "catsmatch'9'"
                    Return "cat = 9"
            End Select

            Return sIn

        End Function

        Private Function SearchString(ByVal xVal As String) As String

            Dim sExtra As String = ""
            Dim sExtra2 As String = ""

            xVal = Replace(xVal, "  ", " ").Replace("  ", " ")
            xVal = Replace$(xVal, "'", "''").Replace("-"c, " ")
            xVal = xVal.Trim

            If Len(xVal) = 0 Then Return vbNullString

            If GebruikFilter.IsEnabled And GebruikFilter.IsChecked And GebruikFilter.Visibility = Windows.Visibility.Visible Then

                Dim TheFil As FilterCat
                TheFil = GetFilter(CType(GebruikFilter.Content, String))

                If Not TheFil Is Nothing Then

                    Dim tQuery As String = RewriteQuery(TheFil.Query)

                    If IsSearchQuery(tQuery) Then

                        sExtra = tQuery & " AND docid IN (SELECT docid FROM search WHERE "
                        sExtra2 = ")"

                    Else

                        GebruikFilter.IsChecked = False

                    End If

                Else

                    GebruikFilter.IsChecked = False

                End If

            End If

            If RadioButton5.IsEnabled And RadioButton5.IsChecked Then
                If Not (Uitgebreid.IsEnabled And Uitgebreid.IsChecked) Then
                    Return sExtra & "sender MATCH '" & xVal.Replace(" ", "").ToLower & "'" & sExtra2
                Else
                    Return sExtra & "sender MATCH '" & xVal.Replace(" ", "").ToLower & "*'" & sExtra2
                End If
            End If

            If RadioButton6.IsEnabled And RadioButton6.IsChecked Then

                Dim Sq As String = ""

                If Not (Uitgebreid.IsEnabled And Uitgebreid.IsChecked) Then

                    If xVal = StripNonAlphaNumericCharacters(xVal) Then
                        Sq = (sExtra & "subject MATCH '" & xVal.ToLower & "'" & sExtra2)
                    Else
                        Sq = (sExtra & "subject MATCH '" & Chr(34) & xVal.ToLower.Replace(Chr(34), "") & Chr(34) & "'" & sExtra2)
                    End If

                Else

                    Sq = (sExtra & "subject MATCH '" & xVal.ToLower.Replace(" ", "* ") & "*" & "'" & sExtra2)

                End If

                Return Sq

            End If

            If RadioButton7.IsEnabled And RadioButton7.IsChecked Then
                If Not (Uitgebreid.IsEnabled And Uitgebreid.IsChecked) Then
                    Return sExtra & "tag MATCH '" & xVal.Replace(" ", "").ToLower & "'" & sExtra2
                Else
                    Return sExtra & "tag MATCH '" & xVal.Replace(" ", "").ToLower & "*'" & sExtra2
                End If
            End If

            Return vbNullString

        End Function

        Private Function DoSearch() As Boolean

            Try

                SearchBox.IsDropDownOpen = False
                SearchText.SelectionStart = SearchText.Text.Length

                Dim sVal As String
                sVal = Trim$(SearchBox.Text)

                If Len(sVal) = 0 Then
                    NoFilter()
                    Return True
                End If

                If My.Settings.GoogleSuggest Then HistoryDB.SaveHistory(sVal)

                ClearFilter()

                SetFilter(SearchM & sVal, SearchString(sVal), "Zoeken naar: " & sVal & "..", True, True, True)

                Return True

            Catch ex As Exception
                Foutje("DoSearch: " & ex.Message)
                Return False

            End Try

        End Function

        Private Function SetFilter(ByVal sName As String, ByVal sFilter As String, ByVal WaitString As String, ByVal bResetCount As Boolean, ByVal bResetScroll As Boolean, ByVal bBlockUI As Boolean) As Boolean

            Dim lPos As Integer = 0

            Try

                lPos = 1

                Dim PerfTimer As New Stopwatch
                PerfTimer.Start()

                If bBlockUI Then DoWait(WaitString, True)

                Spots.SelectedIndex = -1

                lPos = 2

                Dim ScrollViewer As ScrollViewer = CType(EnumVisual(Spots), ScrollViewer)

                If Not ScrollViewer Is Nothing Then
                    If ScrollViewer.VerticalOffset > 0 Then
                        ScrollViewer.ScrollToTop()
                        ScrollViewer.UpdateLayout()
                    End If
                End If

                lPos = 3

                If bResetCount Then SpotSource.ResetCount()

                lPos = 4

                SpotSource.QueryName = sName
                SpotSource.RowFilter = sFilter

                lPos = 5
                IcStopWait = bBlockUI

                Dim TL As DataVirtualization.VirtualList(Of SpotRow) = DirectCast(Spots.ItemsSource, DataVirtualization.VirtualList(Of SpotRow))

                lPos = 6

                TL.Clear()

                lPos = 7

                If (Spots.Items.Count = 0) Or (TabControl1.SelectedIndex <> 0) Then
                    If bBlockUI Then
                        IcStopWait = False
                        DoFinish()
                    End If
                End If

                lPos = 8

                PerfTimer.Stop()
                Debug.Print("LOAD TIME: " & PerfTimer.ElapsedMilliseconds)

                Return True

            Catch ex2 As NullReferenceException

                If lPos = 6 Then

                    Me.Close()
                    Return False

                End If

                Foutje("SetFilter: " & ex2.Message & " (" & CStr(lPos) & ")")
                Return False

            Catch ex As Exception

                Foutje("SetFilter: " & ex.Message & " (" & CStr(lPos) & ")")
                Return False

            End Try

        End Function

        Private Sub ReloadFilter(ByVal bResetCount As Boolean, ByVal bResetScroll As Boolean, Optional ByVal WaitString As String = "Overzicht opbouwen...", Optional ByVal BlockUI As Boolean = True)

            SetFilter(SpotSource.QueryName, SpotSource.RowFilter, WaitString, bResetCount, bResetScroll, BlockUI)

        End Sub

        Private Sub StopWait(ByVal bUpdateStatus As Boolean)

            Dim sIcon As String = "filter.ico"

            If bUpdateStatus Then UpdateStatus()

            If SpotSource.RowFilter = sModule.DefaultFilter Or Len(SpotSource.RowFilter) = 0 Then sIcon = "nofilter.ico"
            If SpotSource.QueryName.StartsWith(SearchM) Then sIcon = "search.ico"

            ChangeName(SpotSource.QueryName, sIcon)
            FirstTab(SpotSource.QueryName, sIcon)

            EnableBijwerken()
            EndWait(bUpdateStatus)

        End Sub

        Private Function NewMsg(ByVal RowsAdded As Integer) As String

            If RowsAdded < 1 Then
                Return "Geen nieuwe spots gevonden"
            Else
                If RowsAdded = 1 Then
                    Return Replace$((RowsAdded).ToString("#,#"), ",", ".") & " nieuwe spot gevonden"
                Else
                    Return Replace$((RowsAdded).ToString("#,#"), ",", ".") & " nieuwe spots gevonden"
                End If
            End If

        End Function

        Private Sub SearchBox_ContextMenuOpening(ByVal sender As Object, ByVal e As System.Windows.Controls.ContextMenuEventArgs) Handles SearchBox.ContextMenuOpening

            CancelSuggest()

            SugTonen.IsChecked = My.Settings.GoogleSuggest

        End Sub

        Private Sub SearchBox_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles SearchBox.KeyUp

            If e.Key = Key.Enter Then
                CancelSuggest()
                If Bijwerken.IsEnabled Then
                    DoSearch()
                    Exit Sub
                End If
            End If

            If e.Key = Key.End Then Exit Sub
            If e.Key = Key.Home Then Exit Sub
            If e.Key = Key.Up Then Exit Sub
            If e.Key = Key.Down Then Exit Sub
            If e.Key = Key.Left Then Exit Sub
            If e.Key = Key.Right Then Exit Sub
            'If e.Key = Key.Delete Then Exit Sub
            'If e.Key = Key.Back Then Exit Sub

            If RadioButton6.IsChecked Then
                QueryGoogle(SearchText.Text)
            End If


        End Sub

        Private Sub QueryGoogle(ByVal sSearch As String)

            Try

                If sSearch.Length < 2 Then
                    If SearchBox.IsDropDownOpen Then SearchBox.IsDropDownOpen = False
                    Exit Sub
                End If

                Static LastSearch As String = ""

                If Len(LastSearch) > 0 Then
                    If LastSearch = sSearch And SearchBox.IsDropDownOpen Then Exit Sub
                End If

                LastSearch = sSearch

                If My.Settings.GoogleSuggest Then

                    Try

                        CancelSuggest()
                        UpdateSuggestions()

                        Dim sU As String = "http://www.google.nl/complete/search?hl=nl&output=toolbar&q=" & HtmlEncode(sSearch)
                        Suggest.DownloadStringAsync(New Uri(sU))

                    Catch
                    End Try

                End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub UpdateSuggestions()

            Try

                Dim s1 As String = SearchBox.Text
                Dim l1 As Integer = SearchText.SelectionStart
                Dim l2 As Integer = SearchText.SelectionLength

                SearchBox.Items.Clear()

                If s1 <> SearchBox.Text Then
                    SearchText.Text = s1
                    SearchText.SelectionStart = l1
                    SearchText.SelectionLength = l2
                End If

                If SuggestList.Count = 0 Then

                    If SearchBox.IsDropDownOpen Then
                        SearchBox.IsDropDownOpen = False
                    End If

                Else

                    For Each JH As String In SuggestList
                        SearchBox.Items.Add(JH)
                    Next

                    If Not SearchBox.IsDropDownOpen Then
                        If My.Settings.GoogleSuggest Then

                            SearchBox.IsDropDownOpen = True
                            SearchText.SelectionStart = l1
                            SearchText.SelectionLength = l2

                            If s1 <> SearchBox.Text Then
                                SearchText.Text = s1
                            End If

                        End If
                    End If

                End If

            Catch
            End Try

        End Sub

        Public Sub StatusChanged(ByVal sMsg As String, ByVal sValue As Long)

            Try

                If Len(sMsg) > 0 Then
                    If StatusMsg.Text <> sMsg Then StatusMsg.Text = sMsg
                End If

                If sValue = 0 Then

                    If Not Me.TaskbarItemInfo Is Nothing Then
                        Me.TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.Indeterminate
                    End If

                Else

                    If sValue > 0 Then

                        If Not Me.TaskbarItemInfo Is Nothing Then
                            If Me.TaskbarItemInfo.ProgressValue <> sValue / 100 Then
                                Me.TaskbarItemInfo.ProgressValue = sValue / 100
                            End If
                            If Me.TaskbarItemInfo.ProgressState <> Shell.TaskbarItemProgressState.Normal Then
                                Me.TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.Normal
                            End If
                        End If

                    Else

                        If Not Me.TaskbarItemInfo Is Nothing Then
                            If Me.TaskbarItemInfo.ProgressState <> Shell.TaskbarItemProgressState.None Then
                                Me.TaskbarItemInfo.ProgressState = Shell.TaskbarItemProgressState.None
                            End If
                        End If

                    End If

                End If

            Catch ex As Exception

                Static ShowOnce As Boolean = False

                If Not ShowOnce Then
                    ShowOnce = True
                    Foutje("StatusChanged:: " & ex.Message)
                End If

            End Try

        End Sub

        Private Sub RefreshSearch()

            If Not IsSearching() Then Exit Sub

            Dim dk As TabItem
            If TabControl1 Is Nothing Then Exit Sub
            If TabControl1.Items.Count = 0 Then Exit Sub
            dk = CType(TabControl1.Items(0), TabItem)

            If GetHeader(dk.Header).Trim = SearchM & Trim$(SearchBox.Text) Then ' Ugly hack
                DoSearch() 'Repeat search 
            End If

        End Sub

        Private Function IsSearching() As Boolean

            Dim dk As TabItem
            If TabControl1 Is Nothing Then Return False
            If TabControl1.Items.Count = 0 Then Return False
            dk = CType(TabControl1.Items(0), TabItem)

            'hack
            Return GetHeader(dk.Header).StartsWith(SearchM)

        End Function

        Private Function TabSearchText() As String

            Dim dk As TabItem
            If TabControl1 Is Nothing Then Return ""
            If TabControl1.Items.Count = 0 Then Return ""
            dk = CType(TabControl1.Items(0), TabItem)

            If Not GetHeader(dk.Header).StartsWith(SearchM) Then Return ""

            'hack
            Return GetHeader(dk.Header).Substring(Len(SearchM))

        End Function

        Private Function FilterSelected() As Boolean

            If FilterList Is Nothing Then Return False
            If FilterList.SelectedItems Is Nothing Then Return False

            Return (FilterList.SelectedItems.Count > 0)

        End Function

        Private Sub GebruikFilter_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles GebruikFilter.Click

            RefreshSearch()

        End Sub

        Private Sub Uitgebreid_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Uitgebreid.Click

            My.Settings.AdvancedSearch = CBool(Uitgebreid.IsChecked)
            My.Settings.Save()

            RefreshSearch()

        End Sub

        Private Sub SpotToevoegen()

            OpenAdd()

        End Sub

        Private Function SelecteerProvider() As Boolean

            Try

                Dim KL As New ProviderSelectie()

                KL.Owner = Me
                KL.ShowDialog()

                Return (Len(Trim$(ServersDB.oDown.Server)) > 0) And KL.bSuc

            Catch
                Foutje("SelecteerProvider::" & Err.Description)
            End Try

            Return False

        End Function

        Public Function IsCloseable(ByRef sTab As TabItem) As Boolean

            If sTab Is Nothing Then Return False
            Return TypeOf sTab Is CloseableTabItem

        End Function

        Public Function IsDownload(ByRef sTab As TabItem) As Boolean

            If sTab Is Nothing Then Return False
            If IsCloseable(sTab) Then Return False

            Try
                If sTab.Tag Is Nothing Then Return False
                Return CStr(sTab.Tag) = "DownloadsTab"
            Catch
                Return False
            End Try

        End Function

        Private Sub TabControl1_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles TabControl1.SelectionChanged

            Dim DoLoad As Boolean = False

            If TabControl1.SelectedIndex = LastTab Then
                Exit Sub
            End If

            If TabControl1.SelectedIndex = 0 Then
                TabControl1.Focus()
                Spots.Focus()
            End If

            LastTab = TabControl1.SelectedIndex

            If IsDownload(CType(TabControl1.SelectedItem, TabItem)) Then
                StartSABDelayed()
                Downloads.Focus()
            End If

            If IsCloseable(CType(TabControl1.SelectedItem, TabItem)) Then
                Dim XX As CloseableTabItem = CType(TabControl1.SelectedItem, CloseableTabItem)

                If TypeOf XX.Tag Is UrlInfo Then
                    Dim Sx As UrlInfo = CType(XX.Tag, UrlInfo)

                    If Not Sx.TabLoaded Then

                        OpenURL(Sx.URL, Sx.Title, XX)
                        Sx.TabLoaded = True

                    End If

                End If

                If TypeOf XX.Tag Is SpotInfo Then
                    Dim Xs As SpotInfo
                    Xs = CType(XX.Tag, SpotInfo)

                    If Not Xs.TabLoaded Then

                        DoLoad = True
                        Xs.TabLoaded = True

                    End If

                    Me.Title = GetHeader(XX.Header) & " - " & Spotname
                    TrayNotify.Text = Microsoft.VisualBasic.Left(Me.Title, 63)

                    If DoLoad Then
                        OpenSpot(-1, MakeMsg(Xs.Spot.MessageID), XX)
                    End If

                    Exit Sub

                End If

            End If

            If Me.Title <> Spotname Then
                Me.Title = Spotname
                TrayNotify.Text = Me.Title
            End If

        End Sub

        Private Sub StartSABDelayed()

            If Not SAB.SabStarted Then

                If SabStartTimer Is Nothing Then
                    SabStartTimer = New DispatcherTimer(DispatcherPriority.Background)
                    SabStartTimer.Interval = TimeSpan.FromMilliseconds(50)
                    AddHandler SabStartTimer.Tick, AddressOf SabDelayed
                    SabStartTimer.Start()
                End If

            Else

                SAB.QuickRefresh()  ' Refresh queue

            End If

        End Sub

        Friend Sub DisplayTooltip(ByVal sTooltip As String)

            If My.Settings.SystemTray Then

                If (TrayNotify.Visible) Then

                    TrayNotify.ShowBalloonTip(1000, Spotname, sTooltip, System.Windows.Forms.ToolTipIcon.Info)

                End If

            End If

        End Sub

        Private Function GUIStartSab() As Boolean

            Dim XX As New Status

            If SAB.SabStarted Then Return True

            AddHandler XX.StatusChanged, AddressOf StatusChanged

            XX.Owner = Me

            XX.Title = "Ogenblik geduld..."
            XX.TheSab = SAB
            XX.ProgressChanged("Downloads starten...", 0)

            StatusChanged(XX.Title, 0)

            DoBlur()

            Dim InnerWait As Cursor = Me.Cursor

            Me.Cursor = System.Windows.Input.Cursors.Wait
            Mouse.OverrideCursor = Cursors.Wait

            XX.ShowDialog()

            Me.Cursor = InnerWait
            Mouse.OverrideCursor = Nothing

            UpdateStatus()
            ClearBlur()

            Dim bRet As Boolean = Not (XX.TheSab Is Nothing)

            XX = Nothing

            Return bRet

        End Function

        Private Sub SabDelayed(ByVal sender As Object, ByVal e As System.EventArgs)

            Dim zErr As String = ""

            SabStartTimer.Stop()
            SabStartTimer.IsEnabled = False
            SabStartTimer = Nothing

            GUIStartSab()

        End Sub

        Private Sub Downloads_ContextMenuOpening(ByVal sender As Object, ByVal e As System.Windows.Controls.ContextMenuEventArgs) Handles Downloads.ContextMenuOpening

            Dim xx As SabItem
            Dim ZX As ContextMenu

            e.Handled = True

            If Downloads.SelectedItems.Count = 0 Then Exit Sub

            Try

                If ((TypeOf e.OriginalSource Is TextBlock)) Then
                    Dim zx1 As TextBlock = CType(e.OriginalSource, TextBlock)
                    If Not TypeOf zx1.Parent Is DataGridCell Then
                        Exit Sub
                    End If
                End If

                If ((TypeOf e.OriginalSource Is Border)) Then
                    Dim zx2 As Border = CType(e.OriginalSource, Border)
                    If Not TypeOf zx2.Parent Is DataGridCell Then
                        Exit Sub
                    End If
                End If

                If Not ((TypeOf e.OriginalSource Is TextBlock) Or (TypeOf e.OriginalSource Is Border)) Then Exit Sub

            Catch
                Exit Sub
            End Try

            If Downloads.SelectedItems.Count > 1 Then

                Dim ZL As New List(Of SabItem)

                Try

                    For Each ZI As SabItem In Downloads.SelectedItems
                        ZL.Add(ZI)
                    Next

                Catch
                    Exit Sub
                End Try

                ZX = SAB.GetMultiMenu(ZL)

            Else

                Try
                    xx = CType(Downloads.SelectedItem, SabItem)
                Catch
                    Exit Sub
                End Try

                ZX = SAB.GetMenu(xx)

            End If

            Dim fe As FrameworkElement = CType(e.Source, FrameworkElement)
            fe.ContextMenu = ZX
            fe.ContextMenu.IsOpen = True
            fe.ContextMenu = Nothing

        End Sub

        Private Sub FilterList_PreviewMouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles FilterList.PreviewMouseDown

            Dim dep As DependencyObject
            dep = CType(e.OriginalSource, DependencyObject)

            If TypeOf dep Is ScrollViewer Then Exit Sub

            If Not TypeOf e.Source Is ListBox Then
                NoFilter()
                e.Handled = True
                Exit Sub
            End If

            Dim X As ListBoxItem
            X = CType(ItemsControl.ContainerFromElement(FilterList, dep), ListBoxItem)

            If e.LeftButton <> MouseButtonState.Pressed Then Exit Sub

            If X Is Nothing Then
                NoFilter()
                e.Handled = True
            Else
                If CStr(X.Content.Name) = LastSel Or Len(Trim$(CStr(X.Content.Name))) = 0 Then
                    NoFilter()
                    e.Handled = True
                End If
            End If

        End Sub

        Private Sub Afsluiten_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Afsluiten.Click

            Me.Close()
            Exit Sub

        End Sub

        Private Sub Bijwerken_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Bijwerken.Click

            If Not Bijwerken.IsEnabled Then Exit Sub

            If (Len(Trim$(ServersDB.oDown.Server)) = 0) Then
                Provider_Click(Nothing, Nothing)
                Exit Sub
            End If

            Dim zErr As String = ""

            If Not DoUpdate(zErr) Then
                If (Len(zErr) > 0) And (zErr <> CancelMSG) Then
                    Foutje(zErr, "Fout")
                End If
            End If

        End Sub

        Private Sub Toevoegen_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Toevoegen.Click

            SpotToevoegen()
            Exit Sub

        End Sub

        Private Sub Provider_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Provider.Click

            Dim zErr As String = ""
            Dim PrevDB As String = SafeName(ServersDB.oDown.Server)

            Try

                DoBlur()

                If SelecteerProvider() Then

                    If SAB.SabStarted Then

                        If Not SAB.UpdateSetting(GetServer(ServerType.Download), zErr) Then Foutje(zErr)

                    End If

                    If PrevDB <> SafeName(ServersDB.oDown.Server) Then

                        If (Len(Trim$(ServersDB.oDown.Server)) > 0) Then

                            If Not OpenDB(zErr) Then

                                Foutje(zErr)
                                Me.Close()
                                Exit Sub

                            End If

                            NoFilter(True)

                            Dim nEmpty(10) As Integer
                            UpdateNewCats(nEmpty)

                            Me.UpdateLayout()

                            If Not DoUpdate(zErr) Then
                                If (Len(zErr) > 0) And (zErr <> CancelMSG) Then
                                    Foutje(zErr, "Fout")
                                End If
                            End If

                        End If
                    End If

                End If

                ClearBlur()

            Catch ex As Exception

                Foutje("Provider_Click: " & ex.Message)
                Me.Close()
                Exit Sub

            End Try

        End Sub

        Private Sub Over_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Over.Click

            OpenAbout()

        End Sub

        Private Sub Map_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Map.Click

            If SAB.External Then
                Exit Sub
            End If

            If SAB.SabStarted Then
                If SAB.HasQueue Then
                    MsgBox("Je kunt deze map niet wijzigen terwijl er nog downloads actief zijn." & vbCrLf & "Zorg eerst dat de wachtrij leeg is, en probeer het dan nog eens.", MsgBoxStyle.Information, "Fout")
                    Exit Sub
                End If
            End If

            Dim ew As New Forms.FolderBrowserDialog

            ew.Description = "Selecteer de map waar je downloads wilt opslaan"
            ew.ShowNewFolderButton = True
            ew.SelectedPath = DownDir()

            DoBlur()
            Dim dr As System.Windows.Forms.DialogResult = ew.ShowDialog()

            If dr = System.Windows.Forms.DialogResult.OK Then

                If DownDir.ToLower <> ew.SelectedPath.ToLower Then

                    My.Settings.DownloadFolder = ew.SelectedPath
                    My.Settings.Save()

                    If SAB.SabStarted Then

                        Dim zErr As String = ""
                        If Not SAB.UpdateFolder(zErr) Then Foutje(zErr)

                    End If

                End If

            End If

            ClearBlur()

        End Sub

        Public Function AskFile(ByVal sFile As String) As String

            Dim ew As New Forms.SaveFileDialog

            ew.AddExtension = True
            ew.AutoUpgradeEnabled = True
            ew.CheckFileExists = False
            ew.CheckPathExists = True
            ew.CreatePrompt = False
            ew.DefaultExt = "nzb"
            ew.Filter = "NZB Bestanden (*.nzb)|*.nzb"
            ew.FilterIndex = 1

            If Len(My.Settings.LastFolder.Trim) > 0 Then
                ew.InitialDirectory = My.Settings.LastFolder
            Else
                ew.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            End If

            ew.OverwritePrompt = True
            ew.RestoreDirectory = True
            ew.Title = "NZB Opslaan"
            ew.FileName = MakeFilename(sFile)

            Dim dr As System.Windows.Forms.DialogResult = ew.ShowDialog()

            If dr <> System.Windows.Forms.DialogResult.OK Then Return vbNullString
            Return ew.FileName

        End Function

        Private Sub SAB_ProgressChanged(ByVal lVal As Integer) Handles SAB.ProgressChanged

            StatusChanged("", lVal)

        End Sub

        Private Sub ZoekPlaatje_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ZoekPlaatje.MouseDown

            If Not Bijwerken.IsEnabled Then Exit Sub
            If Not ZoekPlaatje.Opacity = 1 Then Exit Sub

            If e.LeftButton <> MouseButtonState.Pressed Then Exit Sub

            DoSearch()

        End Sub

        Private Sub RadioButton5_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles RadioButton5.Click

            If RadioButton5.IsChecked Then
                SearchBox.IsDropDownOpen = False
                RefreshSearch()
            End If

        End Sub

        Private Sub RadioButton6_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles RadioButton6.Click

            If RadioButton6.IsChecked Then
                SearchBox.IsDropDownOpen = False
                RefreshSearch()
            End If

        End Sub

        Private Sub RadioButton7_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles RadioButton7.Click

            If RadioButton7.IsChecked Then
                SearchBox.IsDropDownOpen = False
                RefreshSearch()
            End If

        End Sub

        Private Sub ViaSpotnet_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles ViaSpotnet.Click

            If ViaSpotnet.IsChecked Then
                My.Settings.DownloadAction = 1
                My.Settings.Save()
                ShowDownloads(True, True)
            End If

        End Sub

        Private Sub ViaStandaard_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles ViaStandaard.Click

            If ViaStandaard.IsChecked Then
                My.Settings.DownloadAction = 2
                My.Settings.Save()
                ShowDownloads(False, True)
            End If

        End Sub

        Private Sub NZBOpslaan_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles NZBOpslaan.Click

            If NZBOpslaan.IsChecked Then
                My.Settings.DownloadAction = 3
                My.Settings.Save()
                ShowDownloads(False, True)
            End If

        End Sub

        Private Sub DownloadKnop_SubmenuOpened(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles DownloadKnop.SubmenuOpened

            ViaSpotnet.IsChecked = False
            ViaStandaard.IsChecked = False
            NZBOpslaan.IsChecked = False

            Select Case My.Settings.DownloadAction
                Case 0, 1
                    ViaSpotnet.IsChecked = True
                Case 2
                    ViaStandaard.IsChecked = True
                Case 3
                    NZBOpslaan.IsChecked = True
            End Select

        End Sub

        Private Sub AddFilter_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles AddFilter.MouseDown

            If Not btnToevoegen.IsEnabled Then Exit Sub
            If Not AddFilter.IsEnabled Then Exit Sub

            Dim fList As New List(Of FooViewModel)

            Dim TempCat As String
            Dim LF As FooViewModel
            Dim LF2 As FooViewModel
            Dim LF3 As FooViewModel

            Dim SubCats As New Collection

            For lX = 0 To Me.tree.Items.Count - 1
                fList.Add(TryCast(Me.tree.Items(lX), FooViewModel))
            Next

            For Each LF In fList
                If (LF.IsChecked) Is Nothing Then
                    For Each LF2 In LF.Children

                        Dim CntCheck As Long = 0

                        For Each LF3 In LF2.Children
                            If LF3.IsChecked Then CntCheck += 1
                        Next

                        If (CntCheck < LF2.Children.Count) And (CntCheck > 0) Then

                            For Each LF3 In LF2.Children
                                If LF3.IsChecked Then
                                    TempCat = UCase$(LF3.CatLink.Tag)
                                    TempCat = CStr(Val(TempCat.Substring(1)))
                                    If Val(TempCat) > 9 Then
                                        TempCat = UCase$(LF3.CatLink.Tag)
                                    Else
                                        TempCat = UCase$(LF3.CatLink.Tag)
                                        TempCat = TempCat.Substring(0, 1) + "0" + TempCat.Substring(1)
                                    End If
                                    SubCats.Add("X0" & LF.CatLink.Tag & TempCat)
                                End If
                            Next

                        End If
                    Next
                Else
                    If (LF.IsChecked) Then
                        SubCats.Add("X0" & LF.CatLink.Tag)
                    End If
                End If
            Next

            If SubCats.Count = 0 Then
                btnToevoegen.IsEnabled = False
                CheckToevoegen()
                MsgBox("Selecteer minimaal 1 categorie!", MsgBoxStyle.Information)
                Exit Sub
            End If

            If Len(Trim$(Filterbox.Text)) = 0 Then
                btnToevoegen.IsEnabled = False
                CheckToevoegen()
                MsgBox("Geef eerst een naam op voor het filter!", MsgBoxStyle.Information)
                Exit Sub
            End If

            Dim sError As String = Nothing
            Dim zQuery As String = GetFilterString(SubCats)

            If Not SaveFilter(zQuery, Trim$(Filterbox.Text), sError) Then
                Foutje(sError, "Fout")
                Exit Sub
            End If

            Filterbox.Text = ""
            tree.ItemsSource = Spotnet.FooViewModel.CreateFoos

            If Filters.IsExpanded And Expander2.IsExpanded = False Then
                Filters.IsExpanded = False
                stackPanelScrollViewer.ScrollToTop()
            End If

        End Sub

        Private Sub CheckToevoegen()

            If btnToevoegen.IsEnabled Then
                AddFilter.Opacity = 1
                AddFilter.Cursor = Windows.Input.Cursors.Hand
                AddFilter.ToolTip = "Toevoegen..."
            Else
                AddFilter.Opacity = 0.5
                AddFilter.Cursor = Nothing
                AddFilter.ToolTip = Nothing
            End If

        End Sub

        Private Sub Downloads_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Downloads.MouseDoubleClick

            Dim xx As SabItem

            e.Handled = True

            If Downloads.SelectedItems.Count = 0 Then Exit Sub

            Try

                If ((TypeOf e.OriginalSource Is TextBlock)) Then
                    Dim zx1 As TextBlock = CType(e.OriginalSource, TextBlock)
                    If Not TypeOf zx1.Parent Is DataGridCell Then
                        Exit Sub
                    End If
                End If

                If ((TypeOf e.OriginalSource Is Border)) Then
                    Dim zx2 As Border = CType(e.OriginalSource, Border)
                    If Not TypeOf zx2.Parent Is DataGridCell Then
                        Exit Sub
                    End If
                End If

                If Not ((TypeOf e.OriginalSource Is TextBlock) Or (TypeOf e.OriginalSource Is Border)) Then Exit Sub

            Catch
                Exit Sub
            End Try

            If Downloads.SelectedItems.Count = 1 Then

                Try
                    xx = CType(Downloads.SelectedItem, SabItem)
                Catch
                    Exit Sub
                End Try

                SAB.DoOpen(xx)

            End If

        End Sub

        Private Sub UpdateFonts()

            If Me.FontSize <> My.Settings.FontSize Then

                Me.FontSize = My.Settings.FontSize
                ''SB.FontSize = My.Settings.FontSize - 1
                MainMenu.FontSize = My.Settings.FontSize - 1

                For Each ZZ As Spotnet.FilterCat In FilterList.Items
                    ZZ.UpdateFont()
                Next
            End If

        End Sub

        Private Sub Let10_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Let10.Click

            If Let10.IsChecked Then
                My.Settings.FontSize = 10
                My.Settings.Save()
                UpdateFonts()
            End If

        End Sub

        Private Sub Let12_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Let12.Click

            If Let12.IsChecked Then
                My.Settings.FontSize = 12
                My.Settings.Save()
                UpdateFonts()
            End If

        End Sub

        Private Sub Lettergrootte_SubmenuOpened(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Lettergrootte.SubmenuOpened

            Let10.IsChecked = False
            Let12.IsChecked = False
            Let14.IsChecked = False
            Let16.IsChecked = False
            Let18.IsChecked = False
            Let20.IsChecked = False

            Select Case My.Settings.FontSize
                Case 10
                    Let10.IsChecked = True
                Case 12
                    Let12.IsChecked = True
                Case 16
                    Let16.IsChecked = True
                Case 18
                    Let18.IsChecked = True
                Case 20
                    Let20.IsChecked = True
                Case Else
                    Let14.IsChecked = True
            End Select

        End Sub

        Private Sub Let14_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Let14.Click

            If Let14.IsChecked Then
                My.Settings.FontSize = 14
                My.Settings.Save()
                UpdateFonts()
            End If

        End Sub

        Private Sub Let16_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Let16.Click

            If Let16.IsChecked Then
                My.Settings.FontSize = 16
                My.Settings.Save()
                UpdateFonts()
            End If

        End Sub

        Private Sub Let18_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Let18.Click

            If Let18.IsChecked Then
                My.Settings.FontSize = 18
                My.Settings.Save()
                UpdateFonts()
            End If

        End Sub

        Private Sub Let20_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Let20.Click

            If Let20.IsChecked Then
                My.Settings.FontSize = 20
                My.Settings.Save()
                UpdateFonts()
            End If

        End Sub

        Private Sub SetUpdate_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles SetUpdate.Click

            My.Settings.AutoUpdate = SetUpdate.IsChecked
            My.Settings.Save()

        End Sub

        Private Sub Instellingen_SubmenuOpened(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Instellingen.SubmenuOpened

            ShowSug.IsChecked = My.Settings.GoogleSuggest
            ShowCom.IsChecked = My.Settings.ShowComments
            SetTabs.IsChecked = My.Settings.SaveTabs
            SetUpdate.IsChecked = My.Settings.AutoUpdate
            SystemTray.IsChecked = My.Settings.SystemTray

        End Sub

        Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

            Me.Dispatcher.BeginInvoke(New Action(AddressOf InitMain), DispatcherPriority.ApplicationIdle)

        End Sub

        Friend Sub CloseDB()

            If Not SpotSource Is Nothing Then
                SpotSource.Close()
            End If

        End Sub

        Private Sub MainWindow_StateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.StateChanged

            Static DidOnce As Boolean = False

            If My.Settings.SystemTray Then

                If (Me.WindowState = WindowState.Minimized) Then

                    TrayNotify.Visible = True
                    Me.ShowInTaskbar = False

                    If Not DidOnce Then
                        TrayNotify.ShowBalloonTip(500, Spotname, "Klik straks hier om " & Spotname & " weer te openen.", System.Windows.Forms.ToolTipIcon.Info)
                        DidOnce = True
                    End If
                    Exit Sub

                End If

            End If

            If Me.ShowInTaskbar <> True Then Me.ShowInTaskbar = True
            If TrayNotify.Visible <> False Then TrayNotify.Visible = False

            If DownloadsVisible() Then
                If Not SAB Is Nothing Then SAB.QuickRefresh() ' Refresh queue
            End If

        End Sub

        Private Sub TrayNotify_BalloonTipClicked(ByVal sender As Object, ByVal e As System.EventArgs) Handles TrayNotify.BalloonTipClicked

            ''Me.WindowState = WindowState.Normal

        End Sub

        Private Sub SystemTray_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles SystemTray.Click

            My.Settings.SystemTray = SystemTray.IsChecked
            My.Settings.Save()

        End Sub

        Private Sub SearchPopup_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles SearchPopup.Closed

            If Not (SearchBox.SelectedItem Is Nothing) Then
                SearchText.SelectionStart = SearchText.Text.Length
            End If

        End Sub

        Private Sub Filters_Collapsed(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Filters.Collapsed

            If Expander2.IsExpanded = False Then Expander2.IsExpanded = True

        End Sub

        Private Sub Filters_Expanded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Filters.Expanded

            Static FirstExpand As Boolean = True

            If Expander2.IsExpanded Then

                If FirstExpand Then

                    Dim OD As New ObjectDataProvider()

                    OD.MethodName = "CreateFoos"
                    OD.ObjectType = GetType(FooViewModel)

                    tree.DataContext = OD

                End If

                FirstExpand = False
                stackPanelScrollViewer.ScrollToBottom()

            End If

        End Sub

        'Private Sub Filters_Expanded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Filters.Expanded

        '    If Expander2.IsExpanded = True Then Expander2.IsExpanded = False

        'End Sub

        Private Sub Filters_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Filters.MouseDown

            If Filters.IsExpanded Then Exit Sub

            If TypeOf e.OriginalSource Is TextBlock Then
                If e.OriginalSource.text = Filters.Header Then
                    ''If Not Filters.IsExpanded Then Filters.IsExpanded = True
                    Filters.IsExpanded = Not Filters.IsExpanded
                End If
            End If
        End Sub

        Private Sub Expander2_Collapsed(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Expander2.Collapsed

            If Filters.IsExpanded = False Then Filters.IsExpanded = True

        End Sub

        Private Sub Expander2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Expander2.MouseDown

            If Expander2.IsExpanded Then Exit Sub

            If TypeOf e.OriginalSource Is TextBlock Then
                If e.OriginalSource.text = Expander2.Header Then
                    ''If Not Filters.IsExpanded Then Filters.IsExpanded = True
                    Expander2.IsExpanded = Not Expander2.IsExpanded
                End If
            End If

        End Sub

        Private Sub Filterbox_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Filterbox.KeyDown

            If e.Key = Key.Enter Then AddFilter_MouseDown(Nothing, Nothing)

        End Sub

        Private Sub Filterbox_TextChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.TextChangedEventArgs) Handles Filterbox.TextChanged

            btnToevoegen.IsEnabled = TreeChilds() Or (Len(Trim(Filterbox.Text)) > 0)
            CheckToevoegen()

        End Sub

        Private Sub SearchBox_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles SearchBox.Loaded

            SearchPopup = CType(SearchBox.Template.FindName("PART_Popup", SearchBox), Popup)
            SearchText = CType(SearchBox.Template.FindName("PART_EditableTextBox", SearchBox), TextBox)

            SearchText.ContextMenu = SearchBox.ContextMenu

        End Sub

        Private Sub SugTonen_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles SugTonen.Click

            My.Settings.GoogleSuggest = SugTonen.IsChecked
            My.Settings.Save()

        End Sub

        Private Sub MainMenu_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles MainMenu.Loaded

            MainMenu.FontSize = My.Settings.FontSize - 1

        End Sub

        Private Sub Spots_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Spots.MouseDoubleClick

            If e.LeftButton <> MouseButtonState.Pressed Then Exit Sub

            If ((TypeOf e.OriginalSource Is TextBlock)) Then
                Dim zx As TextBlock = CType(e.OriginalSource, TextBlock)
                If TypeOf zx.Parent Is DataGridCell Then
                    e.Handled = True
                    OpenSel()
                    Exit Sub
                End If
            End If

            If ((TypeOf e.OriginalSource Is Border)) Then
                Dim zx As Border = CType(e.OriginalSource, Border)
                If TypeOf zx.Parent Is DataGridCell Then
                    e.Handled = True
                    OpenSel()
                    Exit Sub
                End If
            End If

        End Sub

        Private Sub SetTabs_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles SetTabs.Click

            My.Settings.SaveTabs = SetTabs.IsChecked
            My.Settings.Save()

            If My.Settings.SaveTabs Then SaveTabs("") Else TabDB.ClearTabs()

        End Sub

        Public Sub UpdateTab(ByRef zTab As TabItem, ByVal sTitle As String, ByVal sLoc As String)

            If TypeOf zTab.Tag Is SpotInfo Then

                Dim Zk2 As SpotInfo = CType(zTab.Tag, SpotInfo)

                Zk2.Spot.MessageID = sLoc
                Zk2.Spot.Title = sTitle

            End If

            If TypeOf zTab.Tag Is UrlInfo Then

                Dim Zk As UrlInfo = CType(zTab.Tag, UrlInfo)

                Zk.URL = sLoc
                Zk.Title = sTitle

            End If

            If My.Settings.SaveTabs Then SaveTabs()

        End Sub

        Private Sub Suggest_DownloadStringCompleted(ByVal sender As Object, ByVal e As System.Net.DownloadStringCompletedEventArgs) Handles Suggest.DownloadStringCompleted

            If My.Settings.GoogleSuggest Then

                If Not e.Cancelled Then
                    If e.Error Is Nothing Then

                        Try

                            SuggestList.Clear()

                            Dim xmlDoc As XmlDocument = New XmlDocument()

                            xmlDoc.LoadXml(e.Result)

                            Dim XL As XmlNodeList = xmlDoc.SelectNodes("//CompleteSuggestion")

                            For Each XZ As XmlNode In XL
                                SuggestList.Add(XZ.SelectSingleNode("suggestion/@data").InnerText)
                            Next

                            For Each SD As String In HistoryDB.History
                                If SD.StartsWith(SearchText.Text) Then
                                    SuggestList.Add(SD)
                                End If
                            Next

                            If SearchText.IsFocused Then
                                If SearchBox.IsDropDownOpen Then

                                    UpdateSuggestions()

                                End If
                            End If

                        Catch
                        End Try

                    End If
                End If
            End If

        End Sub

        Private Sub CancelSuggest()

            Try
                Suggest.CancelAsync()
            Catch
            End Try

        End Sub

        Private Sub SearchBox_LostFocus(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles SearchBox.LostFocus

            CancelSuggest()

        End Sub

        Private Sub Spots_LoadingRow(ByVal sender As Object, ByVal e As System.Windows.Controls.DataGridRowEventArgs) Handles Spots.LoadingRow

            Dim IX As DataVirtualization.VirtualListItem(Of SpotRow)

            Try

                IX = CType(e.Row.Item, DataVirtualization.VirtualListItem(Of SpotRow))

                If Not IX.IsLoaded Then

                    Dim ClearCursor As Boolean = Not (Mouse.OverrideCursor Is Cursors.Wait)
                    If ClearCursor Then Mouse.OverrideCursor = Cursors.Wait

                    IX.Load()
                    IX.Reload()

                    If ClearCursor Then Mouse.OverrideCursor = Nothing

                End If

            Catch ex As Exception

                Foutje("LoadingRow: " & ex.Message)
                Exit Sub

            End Try

            Try

                If Not IX.Data Is Nothing Then

                    If (IX.Data.ID > SpotSource.RowNew) And (SpotSource.RowNew > 1) Then
                        If e.Row.FontWeight <> FontWeights.Bold Then
                            e.Row.FontWeight = FontWeights.Bold
                        End If
                    Else
                        If e.Row.FontWeight = FontWeights.Bold Then
                            e.Row.FontWeight = FontWeights.Normal
                        End If
                    End If

                    If WhiteList.Contains(IX.Data.Modulus) Then
                        If Not e.Row.Foreground Is Brushes.DarkGreen Then e.Row.Foreground = Brushes.DarkGreen
                    Else
                        If BlackList.Contains(IX.Data.Modulus) Then
                            If Not e.Row.Foreground Is Brushes.LightGray Then e.Row.Foreground = Brushes.LightGray
                        Else
                            If Not e.Row.Foreground Is Brushes.Black Then e.Row.Foreground = Brushes.Black
                        End If
                    End If

                End If

            Catch
            End Try

        End Sub

        Private Sub LoadHeaderMenu()

            HeaderMenu = New ContextMenu

            HeaderMenu.FontFamily = Me.FontFamily
            HeaderMenu.FontSize = Me.FontSize - 1
            HeaderMenu.FontStyle = Me.FontStyle

            Dim iPos As Integer = 0
            Dim Lx(Spots.Columns.Count - 1) As MenuItem

            For Each XS As DataGridColumn In Spots.Columns
                Dim tk As New MenuItem
                tk.Header = XS.Header
                tk.IsChecked = (XS.Visibility = Windows.Visibility.Visible)
                If Lx(XS.DisplayIndex) Is Nothing Then
                    Lx(XS.DisplayIndex) = tk
                Else
                    Foutje("ColErr")
                End If
            Next

            For zk = 0 To UBound(Lx)
                If Not Lx(zk) Is Nothing Then HeaderMenu.Items.Add(Lx(zk))
            Next

        End Sub

        Private Sub HeaderMenu_PreviewMouseDown(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles HeaderMenu.PreviewMouseDown

            Try
                e.Handled = True
                e.Source.isChecked = Not e.Source.isChecked
                Dim sTxt As String = CStr(e.Source.Header).ToLower
                For Each Xz As DataGridColumn In Spots.Columns
                    If Xz.Header.ToLower = sTxt Then

                        If (Xz.Visibility = Windows.Visibility.Visible) Then

                            For Each Xz2 As DataGridColumn In Spots.Columns
                                If (Xz2.Visibility = Windows.Visibility.Visible) And (Not Xz2 Is Xz) Then
                                    Xz.Visibility = Windows.Visibility.Hidden
                                End If
                            Next

                        Else
                            Xz.Visibility = Windows.Visibility.Visible
                        End If

                        SaveCols()

                    End If
                Next

            Catch ex As Exception

            End Try

        End Sub

        Private Sub SaveCols()

            Dim sCol As String = ""

            For Each Xz As DataGridColumn In Spots.Columns
                If Xz.Visibility = Visibility.Visible Then
                    sCol &= CStr(Xz.DisplayIndex + 1)
                Else
                    sCol &= "0"
                End If
            Next

            sCol += "00000000"
            sCol = Microsoft.VisualBasic.Left(sCol, 10)

            My.Settings.Columns = sCol
            My.Settings.Save()

        End Sub

        Private Function TranslateCol(ByVal zS As String) As String

            Dim sQUery As String = ""

            Select Case CStr(zS).ToLower
                Case "leeftijd"
                    sQUery = "rowid"
                Case "datum"
                    sQUery = "date"
                Case "titel"
                    sQUery = "subject"
                Case "formaat"
                    sQUery = "subcat"
                Case "genre"
                    sQUery = "extcat"
                Case "afzender"
                    sQUery = "sender"
                Case "tag"
                    sQUery = "tag"
                Case "omvang"
                    sQUery = "filesize"
                Case Else
                    sQUery = "rowid"
            End Select

            Return sQUery

        End Function

        Private Function TranslateCol2(ByVal zS As String) As String

            Dim sQUery As String = ""

            Select Case CStr(zS).ToLower

                Case "rowid"
                    sQUery = "leeftijd"

                Case "date"
                    sQUery = "datum"

                Case "title", "subject"
                    sQUery = "titel"

                Case "subcat"
                    sQUery = "formaat"

                Case "extcat"
                    sQUery = "genre"

                Case "sender"
                    sQUery = "afzender"

                Case "tag"
                    sQUery = "tag"

                Case "filesize"
                    sQUery = "omvang"

                Case Else
                    sQUery = zS.ToLower
            End Select

            Return sQUery

        End Function

        Private Sub Spots_Sorting(ByVal sender As Object, ByVal e As System.Windows.Controls.DataGridSortingEventArgs) Handles Spots.Sorting

            Dim lCol As System.Windows.Controls.DataGridTextColumn = CType(e.Column, DataGridTextColumn)

            Static LastCol As String = TranslateCol2(My.Settings.SortColumn)

            If LastCol = CStr(lCol.Header).ToLower Then
                If SpotSource.SortOrder.ToUpper = "DESC" Then
                    SpotSource.SortOrder = "ASC"
                Else
                    SpotSource.SortOrder = "DESC"
                End If
            Else
                SpotSource.SortOrder = "ASC"
            End If

            LastCol = CStr(lCol.Header).ToLower
            SpotSource.SortCol = TranslateCol(LastCol)

            ReloadFilter(False, True, "Sorteren...")

            If SpotSource.SortOrder.ToUpper.Trim = "ASC" Then
                lCol.SortDirection = ComponentModel.ListSortDirection.Ascending
            Else
                lCol.SortDirection = ComponentModel.ListSortDirection.Descending
            End If

            e.Handled = True

        End Sub

        Private Sub SpotMenu_PreviewMouseUp(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles SpotMenu.PreviewMouseUp

            Try

                If e Is Nothing Then Exit Sub
                If e.Source Is Nothing Then Exit Sub
                If e.Source.Tag Is Nothing Then Exit Sub

                Select Case CStr(e.Source.Tag).ToLower
                    Case "report"

                        AddReport(MakeMsg(SpotSource.GetMessageID("spots", SelectedSpot.ID)), SelectedSpot.Titel)

                        Exit Sub

                    Case "delete"

                        DeleteArticle(MakeMsg(SpotSource.GetMessageID("spots", SelectedSpot.ID)), SelectedSpot.Titel)

                        Exit Sub

                    Case "fav"

                        If BlackList.Contains(SelectedSpot.Modulus) Then Exit Sub
                        If Len(SelectedSpot.Modulus) = 0 Then Exit Sub

                        If WhiteList.Contains(SelectedSpot.Modulus) Then

                            RemoveWhite(SelectedSpot.Modulus)

                        Else

                            AddWhite(StripNonAlphaNumericCharacters(SelectedSpot.Afzender), SelectedSpot.Modulus)

                        End If

                        ReloadFilter(False, False)

                        Exit Sub

                    Case "black"

                        If WhiteList.Contains(SelectedSpot.Modulus) Then Exit Sub
                        If Len(SelectedSpot.Modulus) = 0 Then Exit Sub

                        If BlackList.Contains(SelectedSpot.Modulus) Then

                            RemoveBlack(SelectedSpot.Modulus)
                            ReloadFilter(False, False)
                            ShowOnce("Je zult weer spots/reacties van deze afzender gaan ontvangen.", "Zwarte lijst")

                        Else

                            AddBlack(StripNonAlphaNumericCharacters(SelectedSpot.Afzender), SelectedSpot.Modulus)
                            ReloadFilter(False, False)
                            ShowOnce("Je zult geen spots/reacties van deze afzender meer ontvangen.", "Zwarte lijst")

                        End If

                    Case "filter"

                        Dim sError As String = Nothing

                        If Not SaveFilter(SpotSource.RowFilter, TabSearchText, sError) Then
                            Foutje(sError, "Fout")
                        End If

                End Select

            Catch ex As Exception

                Foutje(ex.Message)

            End Try

        End Sub

        Private Sub Downloads_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Downloads.PreviewKeyDown

            If e.Key = Key.Delete Then
                e.Handled = True

                If Downloads.SelectedItems.Count = 0 Then Exit Sub

                For Each XX As SabItem In Downloads.SelectedItems
                    SAB.DeleteDownload(XX)
                Next

                SAB.UpdateSabAsync()
                Exit Sub
            End If

            e.Handled = False

        End Sub

        Private Sub Spot500_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Spot500.Click

            If Spot500.IsChecked Then
                My.Settings.MaxResults = 500
                My.Settings.Save()

                CheckSpotMenu()

                ReloadFilter(True, True)

            End If

        End Sub

        Private Sub Spot5000_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Spot5000.Click

            If (Spot5000.IsChecked) Then
                My.Settings.MaxResults = 5000
                My.Settings.Save()

                CheckSpotMenu()

                ReloadFilter(True, True)

            End If

        End Sub

        Private Sub SpotsAll_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles SpotsAll.Click

            If SpotsAll.IsChecked Then

                My.Settings.MaxResults = -1
                My.Settings.Save()

                CheckSpotMenu()

                ReloadFilter(True, True)

            End If

        End Sub

        Private Sub CheckSpotMenu()

            Spot500.IsChecked = False
            Spot5000.IsChecked = False
            SpotsAll.IsChecked = False

            Select Case My.Settings.MaxResults
                Case 500
                    Spot500.IsChecked = True
                Case 5000
                    Spot5000.IsChecked = True
                Case Else
                    SpotsAll.IsChecked = True
            End Select

        End Sub

        Private Sub UpdateNewCats(ByVal NC() As Integer)

            Try

                Dim lp As Spotnet.FilterCat

                If NC Is Nothing Then ReDim NC(10)

                For Each lp In FilterList.Items

                    Select Case RewriteQuery2(lp.Query).Replace(" ", "").Replace("(", "").Replace(")", "").ToLower
                        Case "cat=1"
                            lp.NewCount = NC(1)
                        Case "cat=2"
                            lp.NewCount = NC(2)
                        Case "cat=3"
                            lp.NewCount = NC(3)
                        Case "cat=4"
                            lp.NewCount = NC(4)
                        Case "cat=5"
                            lp.NewCount = NC(5)
                        Case "cat=6"
                            lp.NewCount = NC(6)
                        Case "cat=9"
                            lp.NewCount = NC(9)
                    End Select

                Next

            Catch ex As Exception
                Foutje("UpdateCats: " & ex.Message)
            End Try

        End Sub

        Private Sub Bij_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Bij.MouseDown

            Bijwerken_Click(Nothing, Nothing)

        End Sub

        Public Shared Function EnumVisual(ByVal myVisual As Visual) As Visual

            Dim TT As Visual = Nothing

            For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(myVisual) - 1

                ' Retrieve child visual at specified index value.
                Dim childVisual As Visual = CType(VisualTreeHelper.GetChild(myVisual, i), Visual)
                If TypeOf (childVisual) Is ScrollViewer Then Return childVisual

                TT = EnumVisual(childVisual)
                If Not TT Is Nothing Then Return TT

            Next i

            Return Nothing

        End Function

        Private Sub HeaderMenu_PreviewMouseUp(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles HeaderMenu.PreviewMouseUp

            SaveCols()

        End Sub

        Private Sub Spots_ColumnReordered(ByVal sender As Object, ByVal e As System.Windows.Controls.DataGridColumnEventArgs) Handles Spots.ColumnReordered

            SaveCols()

        End Sub

        Friend Function HeaderSettings(ByVal bIncludeLast As Boolean, ByVal bUpdateLists As Boolean) As NNTPSettings

            Try

                If bUpdateLists Then
                    UpdateWhitelist()
                    UpdateBlacklist()
                End If

                Dim SS As New NNTPSettings

                SS.BlackList = BlackList()
                SS.WhiteList = WhiteList()
                SS.TrustedKeys = LoadKeys()

                If bIncludeLast Then
                    Try
                        SS.Position = SpotSource.LastPosition("spots")
                    Catch ex As Exception
                        Throw New Exception("HeaderDB: " & ex.Message)
                    End Try
                End If

                SS.GroupName = My.Settings.HeaderGroup
                SS.CheckSignatures = My.Settings.CheckSignatures

                Return SS

            Catch ex As Exception

                Throw New Exception("HeaderSettings: " & ex.Message)

            End Try

        End Function

        Friend Function CommentSettings(ByVal bIncludeLast As Boolean, ByVal bUpdateLists As Boolean) As NNTPSettings

            Try

                If bUpdateLists Then
                    UpdateWhitelist()
                    UpdateBlacklist()
                End If

                Dim SS As New NNTPSettings

                SS.BlackList = BlackList()
                SS.WhiteList = WhiteList()
                SS.TrustedKeys = LoadKeys()

                If bIncludeLast Then

                    Try

                        If FileExists(GetDBFilename("dbc")) Then

                            Dim cDb As New SqlDB

                            If Not cDb.Connect(GetDBFilename("dbc"), True) Then
                                Throw New Exception("Could not read comments database!")
                            End If

                            If Not cDb.ExecuteNonQuery("PRAGMA temp_store = MEMORY;", "") = 0 Then Throw New Exception("PRAGMA temp_store")

                            SS.Position = sModule.LastPosition(cDb, "comments")

                            cDb.Close()

                        Else

                            SS.Position = -1

                        End If

                    Catch ex As Exception

                        Throw New Exception("CommentDB: " & ex.Message)

                    End Try

                End If

                SS.GroupName = My.Settings.ReplyGroup
                SS.CheckSignatures = My.Settings.CheckSignatures

                Return SS

            Catch ex As Exception

                Throw New Exception("CommentSettings: " & ex.Message)

            End Try

        End Function

        Private Sub SpotOpener_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles SpotOpener.DoWork

            Dim xObj() As Object = CType(e.Argument, Object())

            Dim zErr As String = ""
            Dim xRes(4) As Object
            Dim eRes As Boolean = False
            Dim zReturn As String = ""
            Dim xSpot As SpotEx = Nothing
            Dim lID As Long = CLng(xObj(2))
            Dim sM As String = CStr(xObj(0))

            xRes(3) = xObj(1)

            Try

                eRes = SpotClient.Spots.GetSpot(HeaderPhuse, My.Settings.HeaderGroup, lID, sM, zReturn, xSpot, HeaderSettings(False, False), zErr)

                If eRes Then
                    xRes(1) = zReturn
                    xRes(2) = xSpot
                Else
                    xRes(4) = zErr
                End If

            Catch ex As Exception
                eRes = False
                xRes(4) = "GetSpot: " & ex.Message
            End Try

            xRes(0) = eRes
            e.Result = xRes

        End Sub

        Private Sub SpotOpener_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles SpotOpener.RunWorkerCompleted

            Dim xRes() As Object = CType(e.Result, Object())

            Dim x0 As Boolean = CBool(xRes(0))
            Dim x1 As String = CStr(xRes(1))
            Dim x2 As SpotEx = CType(xRes(2), SpotEx)
            Dim x3 As CloseableTabItem = CType(xRes(3), CloseableTabItem)
            Dim x4 As String = CStr(xRes(4))

            SpotOpener = Nothing

            If Not x0 Then

                Foutje(x4, "Fout")

                If Not x3 Is Nothing Then
                    TabControl1.Items.Remove(x3)
                End If

                EndWait(True)
                Exit Sub

            Else

                OpenSpot2(x2, x3)

            End If

        End Sub

        Private Sub MainWindow_Unloaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Unloaded

            Windows.Application.Current.Shutdown()

        End Sub

        Private Sub ShowCom_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles ShowCom.Click

            My.Settings.ShowComments = ShowCom.IsChecked
            My.Settings.Save()

        End Sub

        Private Sub SpotsCount_SubmenuOpened(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles SpotsCount.SubmenuOpened

            CheckSpotMenu()

        End Sub

        Private Sub stackPanelScrollViewer_SizeChanged(ByVal sender As Object, ByVal e As System.Windows.SizeChangedEventArgs) Handles stackPanelScrollViewer.SizeChanged

            Dim t As Thickness = TabControl1.Margin

            If stackPanelScrollViewer.ComputedVerticalScrollBarVisibility = Visibility.Visible Then
                If t.Left <> 7 Then
                    t.Left = 7
                    TabControl1.Margin = t
                End If
            Else
                If t.Left <> 0 Then
                    t.Left = 0
                    TabControl1.Margin = t
                End If
            End If

        End Sub

        Private Sub ShowSug_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles ShowSug.Click

            My.Settings.GoogleSuggest = ShowSug.IsChecked
            My.Settings.Save()

        End Sub

        Private Sub TrayNotify_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TrayNotify.Click

            Try
                Dim tt As Windows.Forms.MouseEventArgs = CType(e, Windows.Forms.MouseEventArgs)
                If tt.Button <> Forms.MouseButtons.Left Then Me.WindowState = WindowState.Normal
            Catch
            End Try

        End Sub

        Private Sub TrayNotify_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles TrayNotify.DoubleClick

            Me.WindowState = WindowState.Normal

        End Sub

        Private Sub sCol_SubmenuOpened(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles sCol.SubmenuOpened

            sCol.Items.Clear()
            LoadHeaderMenu()

            For Each x As MenuItem In HeaderMenu.Items
                Dim zk As New MenuItem
                zk.Header = x.Header
                zk.IsChecked = x.IsChecked
                zk.AddHandler(MenuItem.PreviewMouseDownEvent, New RoutedEventHandler(AddressOf Me.HeaderMenu_PreviewMouseDown), True)
                sCol.Items.Add(zk)
            Next

        End Sub

        Private Sub IC_StatusChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles IC.StatusChanged

            If IC.Status = 2 Then

                If IcStopWait Then

                    IcStopWait = False
                    DoFinish()

                End If

            End If

        End Sub

        Private Sub DoFinish()

            Try

                UpdateSortCol()
                StopWait(True)

            Catch ex As Exception
                Foutje("DoFinish: " & ex.Message)
            End Try

        End Sub

        Private Sub Spots_PreviewMouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Spots.PreviewMouseDown

            If Not Spots.IsFocused Then Spots.Focus()

        End Sub

        Private Sub Downloads_PreviewMouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Downloads.PreviewMouseDown

            If Not Downloads.IsFocused Then Downloads.Focus()

        End Sub

    End Class

End Namespace
