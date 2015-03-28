Imports System.ComponentModel

Namespace Spotnet
    Friend Class FooViewModel
        Implements INotifyPropertyChanged
#Region "Data"

        Public CatLink As SpotCat

        Private _isChecked? As Boolean = False
        Private _isVisible? As Boolean = False
        Private _isExpanded? As Boolean = False
        Private NoCheckbox As Boolean = False

        Private _parent As FooViewModel

#End Region ' Data

#Region "CreateFoos"

        Public Shared Function CreateFoos() As List(Of FooViewModel)

            Dim root As New FooViewModel("Empty")
            Return root.AddCat()

        End Function

        Private Sub New(ByVal name As String)
            Me.Name = name
            Me.Children = New List(Of FooViewModel)()
        End Sub

        Private Sub Initialize()
            For Each child As FooViewModel In Me.Children
                child._parent = Me
                child.Initialize()
            Next child
            If (Me.Children.Count > 0) Then VerifyCheckState()
        End Sub

#End Region ' CreateFoos

#Region "Properties"

        Private privateChildren As List(Of FooViewModel)
        Public Property Children() As List(Of FooViewModel)
            Get
                Return privateChildren
            End Get
            Private Set(ByVal value As List(Of FooViewModel))
                privateChildren = value
            End Set
        End Property

        Private privateIsInitiallySelected As Boolean
        Public Property IsInitiallySelected() As Boolean
            Get
                Return privateIsInitiallySelected
            End Get
            Private Set(ByVal value As Boolean)
                privateIsInitiallySelected = value
            End Set
        End Property

        Private privateName As String
        Public Property Name() As String
            Get
                Return privateName
            End Get
            Private Set(ByVal value As String)
                privateName = value
            End Set
        End Property

#Region "IsChecked"

        ''' <summary>
        ''' Gets/sets the state of the associated UI toggle (ex. CheckBox).
        ''' The return value is calculated based on the check state of all
        ''' child FooViewModels.  Setting this property to true or false
        ''' will set all children to the same check state, and setting it 
        ''' to any value will cause the parent to verify its check state.
        ''' </summary>
        Public Property IsChecked() As Boolean?
            Get
                Return _isChecked
            End Get
            Set(ByVal value? As Boolean)
                Me.SetIsChecked(value, True, True)
            End Set
        End Property

        Private Sub SetIsChecked(ByVal value? As Boolean, ByVal updateChildren As Boolean, ByVal updateParent As Boolean)
            If value = _isChecked Then
                Return
            End If

            _isChecked = value

            If updateChildren AndAlso _isChecked.HasValue Then
                For Each c In Me.Children
                    c.SetIsChecked(_isChecked, True, False)
                Next
            End If

            If updateParent AndAlso _parent IsNot Nothing Then
                _parent.VerifyCheckState()
            End If

            Me.OnPropertyChanged("IsChecked")
        End Sub

        Private Sub VerifyCheckState()
            Dim state? As Boolean = Nothing
            For i As Integer = 0 To Me.Children.Count - 1
                Dim current? As Boolean = Me.Children(i).IsChecked
                If i = 0 Then
                    state = current
                ElseIf Not state.Equals(current) Then
                    state = Nothing
                    Exit For
                End If
            Next i
            Me.SetIsChecked(state, False, True)
        End Sub

#End Region ' IsChecked

#Region "IsExpanded"

        Public Property IsExpanded() As Boolean?
            Get
                Return _isExpanded
            End Get
            Set(ByVal value? As Boolean)
                Me.SetIsExpanded(value, True, True)
            End Set
        End Property

        Private Sub SetIsExpanded(ByVal value? As Boolean, ByVal updateChildren As Boolean, ByVal updateParent As Boolean)
            If value = _isExpanded Then
                Return
            End If

            _isExpanded = value

            Me.OnPropertyChanged("IsExpanded")
        End Sub

#End Region

#Region "IsVisible"

        Public Property IsVisible() As String

            Get
                Return CStr(IIf(NoCheckbox, "Hidden", "Visible"))

            End Get

            Set(ByVal value As String)

                NoCheckbox = Len(value) > 0

            End Set

        End Property

        Public Property TheMargin() As String

            Get
                Return CStr(IIf(NoCheckbox, "0", ""))

            End Get

            Set(ByVal value As String)

            End Set

        End Property


#End Region

#End Region ' Properties

#Region "INotifyPropertyChanged Members"

        Private Sub OnPropertyChanged(ByVal prop As String)
            If Me.PropertyChangedEvent IsNot Nothing Then
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(prop))
            End If
        End Sub

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

#End Region

        Private Function GetSubCat(ByVal sName As String, ByVal sTag As String) As SpotCat

            Dim SubCat As New SpotCat

            SubCat = New SpotCat

            SubCat.Name = sName
            SubCat.Tag = sTag

            Return SubCat

        End Function

        Private Function AddCat() As List(Of FooViewModel)

            Dim xSpotCat As SpotCat
            Dim xSpotCat2 As SpotCat
            Dim xSpotCat3 As SpotCat
            Dim NewCat As New SpotCat
            Dim TheList As New List(Of FooViewModel)

            With NewCat.Children

                .Add(GetSubCat(CatDesc(1), "0"), CatDesc(1))

                With CType(.Item(CatDesc(1)), SpotCat).Children

                    .Add(GetSubCat("Bron", "0b"), "Bron")
                    With CType(.Item("Bron"), SpotCat).Children
                        .Add(GetSubCat("Retail", "b3"))
                        .Add(GetSubCat("Telesync", "b9"))
                        .Add(GetSubCat("R5", "b7"))
                        .Add(GetSubCat("Cam", "b0"))
                    End With

                    .Add(GetSubCat("Taal", "0c"), "Taal")
                    With CType(.Item("Taal"), SpotCat).Children
                        .Add(GetSubCat("Engels gesproken", "c10"))
                        .Add(GetSubCat("Nederlands gesproken", "c11"))
                        .Add(GetSubCat("Duits gesproken", "c12"))
                        .Add(GetSubCat("Frans gesproken", "c13"))
                        .Add(GetSubCat("Spaans gesproken", "c14"))
                        .Add(GetSubCat("Geen ondertitels", "c0"))
                        .Add(GetSubCat("Engels ondertiteld (extern)", "c3"))
                        .Add(GetSubCat("Engels ondertiteld (ingebakken)", "c4"))
                        .Add(GetSubCat("Engels ondertiteld (instelbaar)", "c7"))
                        .Add(GetSubCat("Nederlands ondertiteld (extern)", "c1"))
                        .Add(GetSubCat("Nederlands ondertiteld (ingebakken)", "c2"))
                        .Add(GetSubCat("Nederlands ondertiteld (instelbaar)", "c6"))
                    End With

                    .Add(GetSubCat("Genre", "0d"), "Genre")
                    With CType(.Item("Genre"), SpotCat).Children

                        .Add(GetSubCat("Actie", "d0"))
                        .Add(GetSubCat("Anime", "d29"))
                        .Add(GetSubCat("Animatie", "d2"))
                        .Add(GetSubCat("Aziatisch", "d28"))
                        .Add(GetSubCat("Avontuur", "d1"))
                        .Add(GetSubCat("Cabaret", "d3"))
                        .Add(GetSubCat("Cartoon", "d32"))
                        .Add(GetSubCat("Detective", "c50"))
                        .Add(GetSubCat("Dieren", "c51"))
                        .Add(GetSubCat("Documentaire", "d6"))
                        .Add(GetSubCat("Drama", "d7"))
                        .Add(GetSubCat("Familie", "d8"))
                        .Add(GetSubCat("Fantasie", "d9"))
                        .Add(GetSubCat("Filmhuis", "d10"))
                        .Add(GetSubCat("Historie", "d41"))
                        .Add(GetSubCat("Horror", "d12"))
                        .Add(GetSubCat("Jeugd", "d33"))
                        .Add(GetSubCat("Komedie", "d4"))
                        .Add(GetSubCat("Kort", "d19"))
                        .Add(GetSubCat("Misdaad", "d5"))
                        .Add(GetSubCat("Muziek", "d13"))
                        .Add(GetSubCat("Musical", "d14"))
                        .Add(GetSubCat("Mysterie", "d15"))
                        .Add(GetSubCat("Oorlog", "d21"))
                        .Add(GetSubCat("Romantiek", "d16"))
                        .Add(GetSubCat("Science Fiction", "d17"))
                        .Add(GetSubCat("Sport", "d18"))
                        .Add(GetSubCat("Thriller", "d20"))
                        .Add(GetSubCat("Vrouw", "d46"))
                        .Add(GetSubCat("Western", "d22"))
                        .Add(GetSubCat("Waargebeurd", "d54"))
                    End With

                    .Add(GetSubCat("Formaat", "0a"), "Formaat")
                    With CType(.Item("Formaat"), SpotCat).Children

                        .Add(GetSubCat("MPG", "a2"))
                        .Add(GetSubCat("WMV", "a1"))
                        .Add(GetSubCat("DivX", "a0"))
                        .Add(GetSubCat("DVD5", "a3"))
                        .Add(GetSubCat("DVD9", "a10"))
                        .Add(GetSubCat("Bluray", "a6"))
                        .Add(GetSubCat("x264", "a9"))
                    End With

                End With

                .Add(GetSubCat(CatDesc(6), "5"), CatDesc(6))

                With CType(.Item(CatDesc(6)), SpotCat).Children

                    .Add(GetSubCat("Bron", "5b"), "Bron")
                    With CType(.Item("Bron"), SpotCat).Children
                        .Add(GetSubCat("Retail", "b3"))
                        .Add(GetSubCat("Telesync", "b9"))
                        .Add(GetSubCat("R5", "b7"))
                        .Add(GetSubCat("Cam", "b0"))
                    End With

                    .Add(GetSubCat("Taal", "5c"), "Taal")
                    With CType(.Item("Taal"), SpotCat).Children
                        .Add(GetSubCat("Engels gesproken", "c10"))
                        .Add(GetSubCat("Nederlands gesproken", "c11"))
                        .Add(GetSubCat("Duits gesproken", "c12"))
                        .Add(GetSubCat("Frans gesproken", "c13"))
                        .Add(GetSubCat("Spaans gesproken", "c14"))
                        .Add(GetSubCat("Geen ondertitels", "c0"))
                        .Add(GetSubCat("Engels ondertiteld (extern)", "c3"))
                        .Add(GetSubCat("Engels ondertiteld (ingebakken)", "c4"))
                        .Add(GetSubCat("Engels ondertiteld (instelbaar)", "c7"))
                        .Add(GetSubCat("Nederlands ondertiteld (extern)", "c1"))
                        .Add(GetSubCat("Nederlands ondertiteld (ingebakken)", "c2"))
                        .Add(GetSubCat("Nederlands ondertiteld (instelbaar)", "c6"))
                    End With

                    .Add(GetSubCat("Genre", "5d"), "Genre")
                    With CType(.Item("Genre"), SpotCat).Children

                        .Add(GetSubCat("Actie", "d0"))
                        .Add(GetSubCat("Anime", "d29"))
                        .Add(GetSubCat("Animatie", "d2"))
                        .Add(GetSubCat("Aziatisch", "d28"))
                        .Add(GetSubCat("Avontuur", "d1"))
                        .Add(GetSubCat("Cabaret", "d3"))
                        .Add(GetSubCat("Cartoon", "d32"))
                        .Add(GetSubCat("Detective", "c50"))
                        .Add(GetSubCat("Dieren", "c51"))
                        .Add(GetSubCat("Documentaire", "d6"))
                        .Add(GetSubCat("Drama", "d7"))
                        .Add(GetSubCat("Familie", "d8"))
                        .Add(GetSubCat("Fantasie", "d9"))
                        .Add(GetSubCat("Historie", "d41"))
                        .Add(GetSubCat("Horror", "d12"))
                        .Add(GetSubCat("Jeugd", "d33"))
                        .Add(GetSubCat("Komedie", "d4"))
                        .Add(GetSubCat("Kort", "d19"))
                        .Add(GetSubCat("Misdaad", "d5"))
                        .Add(GetSubCat("Muziek", "d13"))
                        .Add(GetSubCat("Musical", "d14"))
                        .Add(GetSubCat("Mysterie", "d15"))
                        .Add(GetSubCat("Oorlog", "d21"))
                        .Add(GetSubCat("Romantiek", "d16"))
                        .Add(GetSubCat("Science Fiction", "d17"))
                        .Add(GetSubCat("Sport", "d18"))
                        .Add(GetSubCat("Thriller", "d20"))
                        .Add(GetSubCat("Vrouw", "d46"))
                        .Add(GetSubCat("Western", "d22"))
                        .Add(GetSubCat("Waargebeurd", "d54"))
                    End With

                    .Add(GetSubCat("Formaat", "5a"), "Formaat")
                    With CType(.Item("Formaat"), SpotCat).Children

                        .Add(GetSubCat("MPG", "a2"))
                        .Add(GetSubCat("WMV", "a1"))
                        .Add(GetSubCat("DivX", "a0"))
                        .Add(GetSubCat("DVD5", "a3"))
                        .Add(GetSubCat("DVD9", "a10"))
                        .Add(GetSubCat("Bluray", "a6"))
                        .Add(GetSubCat("x264", "a9"))
                    End With

                End With

                .Add(GetSubCat(CatDesc(5), "4"), CatDesc(5))

                With CType(.Item(CatDesc(5)), SpotCat).Children

                    .Add(GetSubCat("Taal", "4c"), "Taal")
                    With CType(.Item("Taal"), SpotCat).Children

                        .Add(GetSubCat("Engels", "c4"))
                        .Add(GetSubCat("Nederlands", "c2"))
                        .Add(GetSubCat("Duits", "c12"))
                        .Add(GetSubCat("Frans", "c13"))
                        .Add(GetSubCat("Spaans", "c14"))

                    End With

                    .Add(GetSubCat("Genre", "4d"), "Genre")
                    With CType(.Item("Genre"), SpotCat).Children

                        .Add(GetSubCat("Avontuur", "d1"))
                        .Add(GetSubCat("Biografie", "d49"))
                        .Add(GetSubCat("Computer", "d35"))
                        .Add(GetSubCat("Cover", "d30"))
                        .Add(GetSubCat("Dagblad", "d43"))
                        .Add(GetSubCat("Detective", "d50"))
                        .Add(GetSubCat("Dieren", "d51"))
                        .Add(GetSubCat("Drama", "d7"))
                        .Add(GetSubCat("Economie", "d34"))
                        .Add(GetSubCat("Fantasie", "d9"))
                        .Add(GetSubCat("Gezondheid", "d40"))
                        .Add(GetSubCat("Handwerk", "d39"))
                        .Add(GetSubCat("Historie", "d41"))
                        .Add(GetSubCat("Hobby", "d36"))
                        .Add(GetSubCat("Jeugd", "d33"))
                        .Add(GetSubCat("Knutselen", "d38"))
                        .Add(GetSubCat("Koken", "d37"))
                        .Add(GetSubCat("Kunst", "d60"))
                        .Add(GetSubCat("Misdaad", "d5"))
                        .Add(GetSubCat("Mysterie", "d15"))
                        .Add(GetSubCat("Non-fictie", "d55"))
                        .Add(GetSubCat("Oorlog", "d21"))
                        .Add(GetSubCat("Poezie", "d57"))
                        .Add(GetSubCat("Psychologie", "d42"))
                        .Add(GetSubCat("Reizen", "d53"))
                        .Add(GetSubCat("Religie", "d47"))
                        .Add(GetSubCat("Roman", "d48"))
                        .Add(GetSubCat("Romantiek", "d16"))
                        .Add(GetSubCat("Science Fiction", "d17"))
                        .Add(GetSubCat("Sport", "d18"))
                        .Add(GetSubCat("Sprookje", "d58"))
                        .Add(GetSubCat("Stripboek", "d31"))
                        .Add(GetSubCat("Studie", "d32"))
                        .Add(GetSubCat("Techniek", "d59"))
                        .Add(GetSubCat("Thriller", "d20"))
                        .Add(GetSubCat("Tijdschrift", "d44"))
                        .Add(GetSubCat("Vrouw", "d46"))
                        .Add(GetSubCat("Waargebeurd", "d54"))
                        .Add(GetSubCat("Wetenschap", "d45"))
                        .Add(GetSubCat("Zakelijk", "d34"))

                    End With

                End With

                .Add(GetSubCat(CatDesc(2), "1"), CatDesc(2))

                With CType(.Item(CatDesc(2)), SpotCat).Children

                    .Add(GetSubCat("Bron", "1b"), "Bron")
                    With CType(.Item("Bron"), SpotCat).Children

                        .Add(GetSubCat("CD", "b0"))
                        .Add(GetSubCat("DVD", "b3"))
                        .Add(GetSubCat("Radio", "b1"))
                        .Add(GetSubCat("Vinyl", "b5"))
                        .Add(GetSubCat("Stream", "b6"))

                    End With

                    .Add(GetSubCat("Type", "1z"), "Type")
                    With CType(.Item("Type"), SpotCat).Children

                        .Add(GetSubCat("Album", "z0"))
                        .Add(GetSubCat("Liveset", "z1"))
                        .Add(GetSubCat("Podcast", "z2"))
                        .Add(GetSubCat("Luisterboek", "z3"))

                    End With

                    .Add(GetSubCat("Genre", "1d"), "Genre")
                    With CType(.Item("Genre"), SpotCat).Children

                        .Add(GetSubCat("Balkan", "d34"))
                        .Add(GetSubCat("Blues", "d0"))
                        .Add(GetSubCat("Cabaret", "d2"))
                        .Add(GetSubCat("Chillout", "d36"))
                        .Add(GetSubCat("Classics", "d24"))
                        .Add(GetSubCat("Compilatie", "d1"))
                        .Add(GetSubCat("Country", "d26"))
                        .Add(GetSubCat("Dance", "d3"))
                        .Add(GetSubCat("Disco", "d23"))
                        .Add(GetSubCat("Diversen", "d4"))
                        .Add(GetSubCat("DnB", "d29"))
                        .Add(GetSubCat("Dubstep", "d27"))
                        .Add(GetSubCat("Electro", "d30"))
                        .Add(GetSubCat("Folk", "d31"))
                        .Add(GetSubCat("Hardstyle", "d5"))
                        .Add(GetSubCat("Hiphop", "d15"))
                        .Add(GetSubCat("Hollands", "d11"))
                        .Add(GetSubCat("Jazz", "d7"))
                        .Add(GetSubCat("Jeugd", "d8"))
                        .Add(GetSubCat("Klassiek", "d9"))
                        .Add(GetSubCat("Latin", "d37"))
                        .Add(GetSubCat("Live", "d38"))
                        .Add(GetSubCat("Metal", "d25"))
                        .Add(GetSubCat("Nederhop", "d28"))
                        .Add(GetSubCat("Pop", "d13"))
                        .Add(GetSubCat("RnB", "d14"))
                        .Add(GetSubCat("Reggae", "d16"))
                        .Add(GetSubCat("Rock", "d18"))
                        .Add(GetSubCat("Soundtrack", "d19"))
                        .Add(GetSubCat("Soul", "d32"))
                        .Add(GetSubCat("Trance", "d33"))
                        .Add(GetSubCat("Techno", "d35"))
                        .Add(GetSubCat("Wereld", "d6"))

                    End With

                    .Add(GetSubCat("Bitrate", "1c"), "Bitrate")
                    With CType(.Item("Bitrate"), SpotCat).Children

                        .Add(GetSubCat("< 96kbit", "c1"))
                        .Add(GetSubCat("96kbit", "c2"))
                        .Add(GetSubCat("128kbit", "c3"))
                        .Add(GetSubCat("160kbit", "c4"))
                        .Add(GetSubCat("192kbit", "c5"))
                        .Add(GetSubCat("256kbit", "c6"))
                        .Add(GetSubCat("320kbit", "c7"))
                        .Add(GetSubCat("Lossless", "c8"))
                        .Add(GetSubCat("Variabel", "c0"))

                    End With

                    .Add(GetSubCat("Formaat", "1a"), "Formaat")

                    With CType(.Item("Formaat"), SpotCat).Children

                        .Add(GetSubCat("MP3", "a0"))
                        .Add(GetSubCat("WMA", "a1"))
                        .Add(GetSubCat("WAV", "a2"))
                        .Add(GetSubCat("OGG", "a3"))
                        .Add(GetSubCat("DTS", "a5"))
                        .Add(GetSubCat("AAC", "a6"))
                        .Add(GetSubCat("APE", "a7"))
                        .Add(GetSubCat("FLAC", "a8"))
                        .Add(GetSubCat("EAC", "a4"))

                    End With

                End With

                .Add(GetSubCat(CatDesc(3), "2"), CatDesc(3))

                With CType(.Item(CatDesc(3)), SpotCat).Children

                    .Add(GetSubCat("Genre", "2c"), "Genre")
                    With CType(.Item("Genre"), SpotCat).Children

                        .Add(GetSubCat("Actie", "c0"))
                        .Add(GetSubCat("Avontuur", "c1"))
                        .Add(GetSubCat("Bordspel", "c13"))
                        .Add(GetSubCat("Educatie", "c15"))
                        .Add(GetSubCat("Jeugd", "c10"))
                        .Add(GetSubCat("Kaart", "c14"))
                        .Add(GetSubCat("Muziek", "c16"))
                        .Add(GetSubCat("Party", "c17"))
                        .Add(GetSubCat("Platform", "c8"))
                        .Add(GetSubCat("Puzzel", "c11"))
                        .Add(GetSubCat("Race", "c5"))
                        .Add(GetSubCat("Rollenspel", "c3"))
                        .Add(GetSubCat("Shooter", "c7"))
                        .Add(GetSubCat("Simulatie", "c4"))
                        .Add(GetSubCat("Sport", "c9"))
                        .Add(GetSubCat("Strategie", "c2"))
                        .Add(GetSubCat("Vliegen", "c6"))

                    End With

                    .Add(GetSubCat("Formaat", "2b"), "Formaat")
                    With CType(.Item("Formaat"), SpotCat).Children

                        .Add(GetSubCat("Rip", "b1"))
                        .Add(GetSubCat("DVD", "b2"))
                        .Add(GetSubCat("DLC", "b3"))
                        .Add(GetSubCat("Patch", "b5"))
                        .Add(GetSubCat("Crack", "b6"))

                    End With

                    .Add(GetSubCat("Platform", "2a"), "Platform")
                    With CType(.Item("Platform"), SpotCat).Children

                        .Add(GetSubCat("Windows", "a0"))
                        .Add(GetSubCat("Linux", "a2"))
                        .Add(GetSubCat("Macintosh", "a1"))
                        .Add(GetSubCat("XBox", "a6"))
                        .Add(GetSubCat("XBox 360", "a7"))
                        .Add(GetSubCat("Gameboy Advance", "a8"))
                        .Add(GetSubCat("Gamecube", "a9"))
                        .Add(GetSubCat("Nintendo DS", "a10"))
                        ''.ADD(GetSubCat("Nintendo 3DS", "a16"))
                        .Add(GetSubCat("Nintendo Wii", "a11"))
                        .Add(GetSubCat("Playstation", "a3"))
                        .Add(GetSubCat("Playstation 2", "a4"))
                        .Add(GetSubCat("Playstation 3", "a12"))
                        .Add(GetSubCat("Playstation Portable", "a5"))
                        .Add(GetSubCat("Windows Phone", "a13"))
                        .Add(GetSubCat("iOs", "a14"))
                        .Add(GetSubCat("Android", "a15"))

                    End With
                End With

                .Add(GetSubCat(CatDesc(4), "3"), CatDesc(4))

                With CType(.Item(CatDesc(4)), SpotCat).Children

                    .Add(GetSubCat("Genre", "3b"), "Genre")
                    With CType(.Item("Genre"), SpotCat).Children
                        .Add(GetSubCat("Audio", "b0"))
                        .Add(GetSubCat("Beveiliging", "b23"))
                        .Add(GetSubCat("Communicatie", "b29"))
                        .Add(GetSubCat("Download", "b15"))
                        .Add(GetSubCat("Educatief", "b26"))
                        .Add(GetSubCat("Foto", "b9"))
                        .Add(GetSubCat("Grafisch", "b2"))
                        .Add(GetSubCat("Internet", "b28"))
                        .Add(GetSubCat("Kantoor", "b27"))
                        .Add(GetSubCat("Ontwikkel", "b30"))
                        .Add(GetSubCat("Spotnet", "b31"))
                        .Add(GetSubCat("Systeem", "b24"))
                        .Add(GetSubCat("Video", "b1"))
                    End With

                    .Add(GetSubCat("Platform", "3a"), "Platform")
                    With CType(.Item("Platform"), SpotCat).Children
                        .Add(GetSubCat("Windows", "a0"))
                        .Add(GetSubCat("Linux", "a2"))
                        .Add(GetSubCat("Macintosh", "a1"))
                        .Add(GetSubCat("Navigatie", "a5"))
                        .Add(GetSubCat("iOs", "a6"))
                        .Add(GetSubCat("Android", "a7"))
                        .Add(GetSubCat("Windows Phone", "a4"))
                    End With

                End With

                .Add(GetSubCat("Erotiek", "8"), "Erotiek")

                With CType(.Item("Erotiek"), SpotCat).Children

                    .Add(GetSubCat("Taal", "8c"), "Taal")

                    With CType(.Item("Taal"), SpotCat).Children
                        .Add(GetSubCat("Engels gesproken", "c10"))
                        .Add(GetSubCat("Nederlands gesproken", "c11"))
                        .Add(GetSubCat("Duits gesproken", "c12"))
                        .Add(GetSubCat("Frans gesproken", "c13"))
                        .Add(GetSubCat("Spaans gesproken", "c14"))
                        .Add(GetSubCat("Geen ondertitels", "c0"))
                        .Add(GetSubCat("Engels ondertiteld (extern)", "c3"))
                        .Add(GetSubCat("Engels ondertiteld (ingebakken)", "c4"))
                        .Add(GetSubCat("Engels ondertiteld (instelbaar)", "c7"))
                        .Add(GetSubCat("Nederlands ondertiteld (extern)", "c1"))
                        .Add(GetSubCat("Nederlands ondertiteld (ingebakken)", "c2"))
                        .Add(GetSubCat("Nederlands ondertiteld (instelbaar)", "c6"))
                    End With

                    .Add(GetSubCat("Genre", "8d"), "Genre")

                    With CType(.Item("Genre"), SpotCat).Children
                        .Add(GetSubCat("Hetero", "d23"))
                        .Add(GetSubCat("Homo", "d24"))
                        .Add(GetSubCat("Lesbo", "d25"))
                        .Add(GetSubCat("Amateur", "d76"))
                        .Add(GetSubCat("BBW", "d84"))
                        .Add(GetSubCat("Bi", "d26"))
                        .Add(GetSubCat("Buiten", "d89"))
                        .Add(GetSubCat("Donker", "d87"))
                        .Add(GetSubCat("Fetisj", "d82"))
                        .Add(GetSubCat("Groep", "d77"))
                        .Add(GetSubCat("Hard", "d86"))
                        .Add(GetSubCat("Hentai", "d88"))
                        .Add(GetSubCat("Jong", "d80"))
                        .Add(GetSubCat("Oud", "d83"))
                        .Add(GetSubCat("POV", "d78"))
                        .Add(GetSubCat("SM", "d85"))
                        .Add(GetSubCat("Soft", "d81"))
                        .Add(GetSubCat("Solo", "d79"))
                    End With

                    .Add(GetSubCat("Formaat", "8a"), "Formaat")

                    With CType(.Item("Formaat"), SpotCat).Children
                        .Add(GetSubCat("MPG", "a2"))
                        .Add(GetSubCat("WMV", "a1"))
                        .Add(GetSubCat("DivX", "a0"))
                        .Add(GetSubCat("DVD5", "a3"))
                        .Add(GetSubCat("DVD9", "a10"))
                        .Add(GetSubCat("Bluray", "a6"))
                        .Add(GetSubCat("x264", "a9"))
                    End With

                End With

            End With

            Dim TempModel As FooViewModel
            Dim TempModel2 As FooViewModel

            For Each xSpotCat In NewCat.Children
                TempModel = New FooViewModel(xSpotCat.Name)
                TempModel.CatLink = xSpotCat
                TempModel.IsExpanded = xSpotCat.Name = CatDesc(1) ' Hack
                With TempModel.Children
                    For Each xSpotCat2 In xSpotCat.Children
                        TempModel2 = New FooViewModel(xSpotCat2.Name)
                        TempModel2.IsVisible = "yes"
                        TempModel2.CatLink = xSpotCat2
                        .Add(TempModel2)
                        With TempModel2.Children
                            For Each xSpotCat3 In xSpotCat2.Children
                                TempModel2 = New FooViewModel(xSpotCat3.Name)
                                TempModel2.CatLink = xSpotCat3
                                .Add(TempModel2)
                            Next
                        End With
                    Next
                End With
                TempModel.Initialize()
                TheList.Add(TempModel)
            Next

            Return TheList

        End Function

    End Class

    Friend Class FilterCat
        Implements INotifyPropertyChanged

        Private Sub OnPropertyChanged(ByVal prop As String)
            If Me.PropertyChangedEvent IsNot Nothing Then
                RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(prop))
            End If
        End Sub

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Public sName As String = ""
        Public sDisplay As String = ""
        Public iNew As Integer = 0
        Public sQuery As String = ""
        Public sMargin As String = ""
        Public Visible As Boolean = True

        Public sImage As String = ""
        Public sSelected As String = ""

        Private bSel As Boolean = False

        Public ReadOnly Property Name As String
            Get
                Return sName
            End Get
        End Property

        Public ReadOnly Property DisplayText As String
            Get

                If Len(Trim(sName)) = 0 Then Return ""

                If iNew > 0 Then
                    Return sName & " (" & iNew & ")"
                End If

                Return sName

            End Get
        End Property

        Public Property NewCount As Integer
            Get
                Return iNew
            End Get
            Set(ByVal value As Integer)
                If iNew <> value Then
                    iNew = value
                    Me.OnPropertyChanged("DisplayText")
                End If
            End Set
        End Property

        Public ReadOnly Property Image As String
            Get

                If bSel And Len(sSelected) > 0 Then
                    Return sSelected
                Else
                    Return sImage
                End If

            End Get
        End Property

        Public ReadOnly Property SelectedImage As String
            Get

                Return sSelected

            End Get
        End Property

        Public ReadOnly Property NormalImage As String
            Get

                Return sImage

            End Get
        End Property

        Public ReadOnly Property Query As String
            Get
                Return sQuery
            End Get
        End Property

        Public ReadOnly Property Fontsize As Integer

            Get
                Return My.Settings.FontSize + 1
            End Get

        End Property

        Public ReadOnly Property Margin As String
            Get
                Return sMargin
            End Get
        End Property

        Public Sub UpdateFont()

            Me.OnPropertyChanged("Fontsize")

        End Sub

        Public Property Selected As Boolean

            Get
                Return bSel
            End Get

            Set(ByVal value As Boolean)

                If bSel <> value Then
                    bSel = value
                    If Len(sSelected) > 0 Then UpdateImage()
                End If

            End Set

        End Property

        Private Sub UpdateImage()

            Me.OnPropertyChanged("Image")

        End Sub

    End Class

End Namespace