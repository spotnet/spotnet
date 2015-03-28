Imports System.Text.RegularExpressions
Imports System.Windows.Threading

Public Class ProviderSelectie

    Public bSuc As Boolean

    Friend Sub New()

        InitializeComponent()

    End Sub

    Private Sub ProviderSelectie_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing

        If Not e.Cancel Then
            Try
                Me.Owner.Activate() 'WPF bug
            Catch
            End Try
        End If

    End Sub

    Private Sub ProviderSelectie_Initialized(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Initialized

        Dim Mnu() As String
        Dim menuEntry As ProviderItem
        Dim ProviderLijst As New Collection
        Dim LP As ProviderItem

        Try

            If Me.FontSize <> My.Settings.FontSize Then
                Me.FontSize = My.Settings.FontSize
            End If

            ProviderLijst.Add("Astraweb#eu.news.astraweb.com#119")
            ProviderLijst.Add("Caiway#news.caiway.nl#119")
            ProviderLijst.Add("Casema#news.casema.nl#119")
            ProviderLijst.Add("Cistron#newsbin.cistron.nl#119")
            ProviderLijst.Add("Dommel#news.dommel.be#119")
            ProviderLijst.Add("Euronet#news.euronet.nl#119")
            ProviderLijst.Add("Eweka#newsreader1.eweka.nl#119")
            ProviderLijst.Add("Giganews#news-europe.giganews.com#119")
            ProviderLijst.Add("Freeler#news.freeler.nl#119")
            ProviderLijst.Add("HetNet#nova.planet.nl#119")
            ProviderLijst.Add("KPN#nova.planet.nl#119")
            ProviderLijst.Add("Lightning#news.lightningusenet.com#119")
            ProviderLijst.Add("NewsXS#reader2.newsxs.nl#119")
            ProviderLijst.Add("Online#news.online.nl#119")
            ProviderLijst.Add("Quicknet#news.quicknet.nl#119")
            ProviderLijst.Add("SpeedXS#news.speedxs.nl#119")
            ProviderLijst.Add("Supernews#news.eu.supernews.com#119")
            ProviderLijst.Add("TweakDSL#news.tweakdsl.nl#119")
            ProviderLijst.Add("Tiscali#news.tiscali.nl#119")
            ProviderLijst.Add("UPC#news.upc.nl#119")
            ProviderLijst.Add("XLned#reader2.xlned.com#119")
            ProviderLijst.Add("Xs4All#newszilla.xs4all.nl#119")
            ProviderLijst.Add("XsNews#reader.xsnews.nl#119")
            ProviderLijst.Add("Ziggo#news.ziggo.nl#119")
            ProviderLijst.Add(" ")
            ProviderLijst.Add("Anders...##119")

            For Each sProv As String In ProviderLijst
                If Len(sProv.Trim) > 0 Then
                    Mnu = Split(sProv, "#")
                    menuEntry = New ProviderItem
                    menuEntry.Address = Mnu(1)
                    menuEntry.Name = Mnu(0)
                    menuEntry.Port = CLng(Mnu(2))
                    ProviderBox.Items.Add(menuEntry)
                Else
                    menuEntry = New ProviderItem
                    ProviderBox.Items.Add(menuEntry)
                End If
            Next

            Dim LZ As ServerInfo = ServersDB.oDown
            Dim LZ9 As ServerInfo = ServersDB.oHeader

            TextBox1.Text = LZ.Server

            If Len(TextBox1.Text) > 0 Then
                For rx = 0 To ProviderBox.Items.Count - 1
                    LP = CType(ProviderBox.Items.Item(rx), ProviderItem)
                    If LP.Address.IndexOf(TextBox1.Text) > -1 Then
                        ProviderBox.SelectedIndex = rx
                        Exit For
                    End If
                Next
            End If

            If (ProviderBox.SelectedIndex) = -1 Then
                If Len(TextBox1.Text) > 0 Then ProviderBox.SelectedIndex = ProviderBox.Items.Count - 1
            End If

            'Herhalen ivm selectedindex event providerbox
            TextBox1.Text = LZ.Server

            Button2.IsEnabled = Len(TextBox1.Text) > 0

            Select Case LZ.Port
                Case 119, 0
                    ComboBox1.SelectedIndex = 0
                Case 443
                    ComboBox1.SelectedIndex = 1
                Case 563
                    ComboBox1.SelectedIndex = 2
                Case Else
                    ComboBox1.Text = CStr(LZ.Port)
            End Select

            ComboBox2.Text = CStr(LZ.Connections + LZ9.Connections)

            Dim KL As New cEncPass

            TextBox2.Text = LZ.Username

            If Len(LZ.Password) > 0 Then
                TextBox3.Password = LZ.Password
            Else
                TextBox3.Password = ""
            End If

        Catch ex As Exception
            Foutje(ex.Message)
            Me.Close()
            Exit Sub
        End Try

    End Sub

    Private Sub EnableVal(ByVal bVal As Boolean)

        Button2.IsEnabled = bVal
        TextBox1.IsEnabled = bVal
        ComboBox1.IsEnabled = bVal
        ComboBox2.IsEnabled = bVal
        TextBox2.IsEnabled = bVal
        TextBox3.IsEnabled = bVal

    End Sub

    Private Sub ProviderBox_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ProviderBox.SelectionChanged

        Dim Kl As ProviderItem = CType(ProviderBox.SelectedItem, ProviderItem)

        TextBox1.Text = CStr(Kl.Address)

        Select Case CLng(Kl.Port)
            Case 119, 0
                ComboBox1.SelectedIndex = 0
            Case 443
                ComboBox1.SelectedIndex = 1
            Case 563
                ComboBox1.SelectedIndex = 2
            Case Else
                ComboBox1.Text = CStr(Kl.Port)
        End Select

        TextBox2.Clear()
        TextBox3.Clear()

        ComboBox2.Text = "4"

        EnableVal(True)

    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button2.Click

        Dim kl As New cEncPass
        Dim sError As String = ""
        Dim NoPassword As Boolean = False

        Button2.IsEnabled = False
        ProviderBox.IsEnabled = False
        EnableVal(False)

        Me.Cursor = Cursors.Wait
        Mouse.OverrideCursor = Cursors.Wait

        Me.UpdateLayout()
        Me.Dispatcher.BeginInvoke(New Action(AddressOf Me.DoButton), DispatcherPriority.Background)

    End Sub

    Private Sub DoButton()

        Dim kl As New cEncPass
        Dim sError As String = ""
        Dim NoPassword As Boolean = False
        Dim CIDownloadServer As String = ""
        Dim CIUploadServer As String = ""
        Dim CIsServer As String = ""

        Dim NC As New ServerInfo
        Dim NU As New ServerInfo
        Dim ND As New ServerInfo
        Dim CustomSSL As Boolean

        If TextBox1.Text.ToLower.Contains(".eweka.") Then
            CIDownloadServer = TextBox1.Text.ToLower
            CIUploadServer = "upload.eweka.nl"
            CIsServer = "textnews.eweka.nl"
        Else
            If TextBox1.Text.ToLower.Contains(".planet.nl") Then
                CIDownloadServer = TextBox1.Text.ToLower
                CIUploadServer = "text.nova.planet.nl"
                CIsServer = "text.nova.planet.nl"
            Else
                If TextBox1.Text.ToLower.Contains(".kpn.nl") Or TextBox1.Text.ToLower.Contains(".kpnplanet.nl") Then
                    CIDownloadServer = TextBox1.Text.ToLower
                    CIUploadServer = "textnews.kpn.nl"
                    CIsServer = "textnews.kpn.nl"
                Else
                    If TextBox1.Text.ToLower.Contains(".caiway.") Then
                        CIDownloadServer = TextBox1.Text.ToLower
                        CIUploadServer = "txtnews.caiway.nl"
                        CIsServer = "txtnews.caiway.nl"
                    Else
                        If TextBox1.Text.ToLower.Contains(".xs4all.") Then
                            CIDownloadServer = TextBox1.Text.ToLower
                            CIUploadServer = "news.xs4all.nl"
                            CIsServer = "news.xs4all.nl"
                        Else
                            CustomSSL = True
                            CIsServer = Trim$(TextBox1.Text).ToLower
                        End If
                    End If
                End If
            End If
        End If

        NC.Server = CIsServer
        NU.Server = sIIF(Len(CIUploadServer) > 0, CIUploadServer, CIsServer)
        ND.Server = sIIF(Len(CIDownloadServer) > 0, CIDownloadServer, CIsServer)

        NC.Port = CType(RemoveStrings(ComboBox1.Text), Integer)
        NU.Port = NC.Port

        ND.Port = NC.Port
        ND.SSL = (ND.Port = 443 Or ND.Port = 563)

        ND.Connections = CType(RemoveStrings(ComboBox2.Text), Integer) - 2
        If ND.Connections < 1 Then ND.Connections = 1

        NC.Connections = 2
        NU.Connections = 1

        If Not NoPassword Then

            NC.Username = TextBox2.Text
            NU.Username = NC.Username
            ND.Username = NC.Username

            NC.Password = TextBox3.Password
            NU.Password = NC.Password
            ND.Password = NC.Password

        End If

        If CustomSSL Then
            NC.SSL = ND.SSL
            NU.SSL = ND.SSL
        Else
            NC.SSL = False
            NU.SSL = False
            NC.Port = 119
            NU.Port = 119
        End If

        ClearPhuses()

        Dim TestPhuse As Phuse.Engine = CreatePhuse(ND)

        If SpotClient.Spots.TestConnection(TestPhuse, My.Settings.HeaderGroup, sError) Then

            If ND.Server <> NC.Server Then
                TestPhuse.Close()
                TestPhuse = CreatePhuse(NC)
            End If

            If SpotClient.Spots.TestConnection(TestPhuse, My.Settings.HeaderGroup, sError) Then

                bSuc = True

                TestPhuse.Close()
                TestPhuse = Nothing

                ClearPhuses()

                ServersDB.oUp = NU
                ServersDB.oDown = ND
                ServersDB.oHeader = NC

                ServersDB.SaveServers()

                Me.Cursor = Nothing
                Mouse.OverrideCursor = Nothing

                Me.Close()

                Exit Sub

            End If

        End If

        TestPhuse.Close()
        TestPhuse = Nothing

        Me.Cursor = Nothing
        Mouse.OverrideCursor = Nothing

        Foutje(sError)
        Button2.IsEnabled = True
        ProviderBox.IsEnabled = True
        EnableVal(True)

    End Sub

    Private Function RemoveStrings(ByVal msg As String) As String

        Dim Number As String = String.Empty
        Dim NegativeValue As String
        Dim value As String = "-"
        Dim returnValue As Boolean

        Dim matches As MatchCollection = Regex.Matches(msg, "[0-9.-]")

        For Each i As Match In matches
            Number &= i.ToString()
        Next

        NegativeValue = Number
        returnValue = NegativeValue.Contains(value)

        If returnValue = False Then
            Number = ""
            Dim matches1 As MatchCollection = Regex.Matches(msg, "[0-9.+]")
            For Each i As Match In matches1
                Number &= i.ToString()
            Next
        End If

        If Number = String.Empty Then Number = "0"
        If String.IsNullOrEmpty(msg) Then Number = "0"

        Return Replace$(Number, "-", "") ' - Symbool weghalen

    End Function

    Private Sub TextBox1_TextChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.TextChangedEventArgs) Handles TextBox1.TextChanged

        Button2.IsEnabled = Len(Trim$(TextBox1.Text)) > 0

    End Sub

End Class
