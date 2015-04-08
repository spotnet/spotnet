Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.Threading
Imports System.Security.Cryptography
Imports System.Windows.Media.Imaging
Imports System.Data.Common

Friend Module sModule

    Friend LastServer As String = ""

    Friend ServersDB As New cServers
    Friend Const SearchM As String = "Zoeken: "
    Friend Const CancelMSG As String = "Geannuleerd"
    Friend Const Spotname As String = "Spotnet"
    Friend Const DefaultFilter As String = "cat < 9"

    Private iWhiteList As HashSet(Of String)
    Private iBlackList As HashSet(Of String)

    Private hPhuse As Phuse.Engine = Nothing
    Private dPhuse As Phuse.Engine = Nothing
    Private uPhuse As Phuse.Engine = Nothing

    Friend ReadOnly EPOCH As Date = New Date(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)

    Private Declare Function SHGetKnownFolderPath Lib "shell32" (ByRef knownFolder As Guid, ByVal flags As UInteger, ByVal htoken As IntPtr, ByRef path As IntPtr) As Integer

    Private Enum convTo
        B = 0
        KB = 1
        MB = 2
        GB = 3  'Enumerations for file size conversions
        TB = 4
        PB = 5
        EB = 6
        ZI = 7
        YI = 8
    End Enum

    Public Function ConvertToTimestamp(ByVal value As DateTime) As Integer

        Dim span As TimeSpan = (value - New DateTime(1970, 1, 1, 0, 0, 0, 0).ToLocalTime)
        Return CInt(span.TotalSeconds)

    End Function

    Public Function GetFileSize(ByVal MyFilePath As String) As Long

        If Not FileExists(MyFilePath) Then Return 0

        Dim MyFile As New FileInfo(MyFilePath)
        Dim FileSize As Long = MyFile.Length
        Return FileSize

    End Function

    Public Sub Wait(ByVal ms As Integer)

        Using wh As New ManualResetEvent(False)
            wh.WaitOne(ms)
        End Using

    End Sub

    Public Function MakeMsg(ByVal sMes As String, Optional ByVal Tag As Boolean = True) As String

        If sMes.Substring(0, 1) = "<" Then
            If Tag Then Return sMes
        Else
            If Not Tag Then Return sMes
        End If

        If Tag Then
            Return "<" & sMes & ">"
        Else
            Return sMes.Substring(1, sMes.Length - 2)
        End If

    End Function

    Private Function AddDirSep(ByVal strPathName As String) As String

        If Microsoft.VisualBasic.Right(Trim(strPathName), 1) <> "\" Then Return Trim(strPathName) & "\"
        Return Trim(strPathName)

    End Function

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


    Public Function DirectoryExists(ByVal sDirName As String) As Boolean

        Try
            Dim dDir As New DirectoryInfo(AddDirSep(sDirName))
            Return dDir.Exists
        Catch
            Return False
        End Try

    End Function

    Public Function FileExists(ByVal sFilename As String) As Boolean

        Try
            If Len(sFilename) = 0 Then Return False
            Dim fFile As New FileInfo(sFilename)
            Return fFile.Exists
        Catch
            Return False
        End Try

    End Function

    Public Function HtmlEncode(ByVal text As String) As String

        If Len(text) = 0 Then Return ""

        Dim sOut As String = System.Net.WebUtility.HtmlEncode(HtmlDecode(text))

        Dim chars() As Char = sOut.ToCharArray
        Dim tOut As New StringBuilder(chars.Length * 2)

        For Each c As Char In chars

            Dim value As Integer = Convert.ToInt32(c)

            If (value > 31) And (value < 127) And (value <> 96) Then
                tOut.Append(c)
            Else
                tOut.Append("&#")
                tOut.Append(value)
                tOut.Append(";")
            End If

        Next

        Return tOut.ToString

    End Function

    Public Function HtmlDecode(ByVal text As String) As String

        If Len(text) = 0 Then Return ""

        Return System.Net.WebUtility.HtmlDecode(text.Replace("&amp;", "&")).Replace(vbLf, "").Replace(vbCr, "").Replace(vbTab, "")

    End Function

    Public Function SafeHref(ByVal Text As String) As String

        Return AddHttp(Text).Replace(Chr(34), "%22").Replace("`", "%60").Replace("'", "%27")

    End Function

    Public Function AddHttp(ByVal Text As String) As String

        If Not HasHttp(Text) And Text.Length > 0 Then Return "http://" & Text
        Return Text

    End Function

    Public Function HasHttp(ByVal Text As String) As Boolean

        If Text.IndexOf(":") > 1 Then

            Select Case Split(Text, ":")(0).ToLower
                Case "http", "https"
                    Return True
            End Select

        End If

        Return False

    End Function

    Public Function URLEncode(ByVal Text As String) As String

        Dim i As Integer
        Dim aCode As Integer
        Dim xURLEncode As String = Text

        For i = Len(xURLEncode) To 1 Step -1

            aCode = Asc(Mid$(xURLEncode, i, 1))

            Select Case aCode

                Case 48 To 57, 65 To 90, 97 To 122
                    ' don't touch alphanumeric chars

                Case 32
                    ' replace space with "+"
                    Mid$(xURLEncode, i, 1) = "+"

                Case 42, 46, 47, 58, 95 ' / : . - _
                    ' sla over

                Case Else

                    ' replace punctuation chars with "%hex"
                    xURLEncode = Microsoft.VisualBasic.Strings.Left(xURLEncode, i - 1) & "%" & Hex$(aCode) & Mid$(xURLEncode, i + 1)

            End Select
        Next

        Return xURLEncode

    End Function

    Public Function SafeName(ByVal Text As String) As String

        Dim i As Integer
        Dim aCode As Integer
        Dim sOut As String = ""
        Dim AllowNumeric As Boolean = False

        Text = Text.Trim
        If IsNumeric(Text.Replace(".", "").Replace(":", "")) Then AllowNumeric = True

        For i = Len(Text) To 1 Step -1
            aCode = Asc(Mid$(Text, i, 1))
            Select Case aCode

                Case 48 To 57
                    ' cijfers niet toestaan (ivm reader13.eweka.nl, etc.)

                    If AllowNumeric Then
                        sOut = Chr(aCode) + sOut
                    End If

                Case 65 To 90, 97 To 122    ' lettters
                    sOut = Chr(aCode) + sOut

                Case 97 To 122 ' lettters
                    sOut = Chr(aCode - 32) + sOut

                Case 46 ' puntje
                    sOut = Chr(aCode) + sOut

                Case Else

                    ' Doe niets

            End Select
        Next

        Return sOut

    End Function

    Public Function URLDecode(ByVal StringToDecode As String) As String

        Dim TempAns As String = ""
        Dim CurChr As Integer = 1

        Do Until CurChr - 1 = Len(StringToDecode)
            Select Case Mid(StringToDecode, CurChr, 1)
                Case "+"
                    TempAns = TempAns & " "
                Case "%"
                    TempAns = TempAns & Chr(CInt(Val("&h" & Mid(StringToDecode, CurChr + 1, 2))))
                    CurChr = CurChr + 2
                Case Else
                    TempAns = TempAns & Mid(StringToDecode, CurChr, 1)
            End Select
            CurChr = CurChr + 1
        Loop

        URLDecode = TempAns

    End Function

    Friend Function CatDesc(ByVal hCat As Byte, Optional ByVal zCat As Byte = 0) As String

        Select Case hCat

            Case 1

                If zCat < 2 Then Return "Films"

                Select Case zCat

                    Case 2
                        Return "Series"

                    Case 3
                        Return "Boeken"

                    Case 4
                        Return "Erotiek"

                End Select

            Case 2

                If zCat < 2 Then Return "Muziek"

                Select Case zCat

                    Case 2
                        Return "Liveset"

                    Case 3
                        Return "Podcast"

                    Case 4
                        Return "Audiobook"

                End Select

            Case 3
                Return "Spellen"

            Case 4
                Return "Applicaties"

            Case 5
                Return "Boeken"

            Case 6
                Return "Series"

            Case 9
                Return "Erotiek"

        End Select

        Return "Fout"

    End Function

    Public Function TranslateCat(ByVal hCat As Long, ByVal sCat As String, Optional ByVal bStrict As Boolean = False) As String

        If sCat.Length < 2 Then Return ""

        Select Case hCat

            Case 2

                Select Case sCat.Substring(0, 1)

                    Case "a"

                        Select Case CInt(sCat.Substring(1))
                            Case 0 : Return "MP3"
                            Case 1 : Return "WMA"
                            Case 2 : Return "WAV"
                            Case 3 : Return "OGG"
                            Case 4 : Return "EAC"
                            Case 5 : Return "DTS"
                            Case 6 : Return "AAC"
                            Case 7 : Return "APE"
                            Case 8 : Return "FLAC"
                            Case Else : Return ""
                        End Select

                    Case "b"

                        Select Case CInt(sCat.Substring(1))
                            Case 0 : Return "CD"
                            Case 1 : Return "Radio"
                            Case 3 : Return "DVD"
                            Case 5 : Return "Vinyl"
                            Case 2 : Return sIIF(bStrict, "", "Compilatie")
                            Case 4 : Return "" ''"Anders"
                            Case 6 : Return "Stream"
                            Case Else : Return ""
                        End Select

                    Case "c"

                        Select Case CInt(sCat.Substring(1))
                            Case 1 : Return "< 96kbit"
                            Case 2 : Return "96kbit"
                            Case 3 : Return "128kbit"
                            Case 4 : Return "160kbit"
                            Case 5 : Return "192kbit"
                            Case 6 : Return "256kbit"
                            Case 7 : Return "320kbit"
                            Case 8 : Return "Lossless"
                            Case 0 : Return "Variabel"
                            Case 9 : Return "" ''"Anders"
                            Case Else : Return ""
                        End Select

                    Case "d"
                        Select Case CInt(sCat.Substring(1))
                            Case 0 : Return "Blues"
                            Case 1 : Return "Compilatie"
                            Case 2 : Return "Cabaret"
                            Case 3 : Return "Dance"
                            Case 4 : Return "Diversen"
                            Case 5 : Return "Hardstyle"
                            Case 6 : Return "Wereld"
                            Case 7 : Return "Jazz"
                            Case 8 : Return "Jeugd"
                            Case 9 : Return "Klassiek"
                            Case 10 : Return sIIF(bStrict, "", "Kleinkunst")
                            Case 11 : Return "Hollands"
                            Case 12 : Return sIIF(bStrict, "", "New Age")
                            Case 13 : Return "Pop"
                            Case 14 : Return "RnB"
                            Case 15 : Return "Hiphop"
                            Case 16 : Return "Reggae"
                            Case 17 : Return "Religieus"
                            Case 18 : Return "Rock"
                            Case 19 : Return "Soundtrack"
                            Case 20 : Return "" ''"Anders"
                            Case 21 : Return sIIF(bStrict, "", "Hardstyle")
                            Case 22 : Return sIIF(bStrict, "", "Aziatisch")
                            Case 23 : Return "Disco"
                            Case 24 : Return "Classics"
                            Case 25 : Return "Metal"
                            Case 26 : Return "Country"
                            Case 27 : Return "Dubstep"
                            Case 28 : Return sIIF(bStrict, "", "Nederhop")
                            Case 29 : Return "DnB"
                            Case 30 : Return "Electro"
                            Case 31 : Return "Folk"
                            Case 32 : Return "Soul"
                            Case 33 : Return "Trance"
                            Case 34 : Return "Balkan"
                            Case 35 : Return "Techno"
                            Case 36 : Return "Ambient"
                            Case 37 : Return "Latin"
                            Case 38 : Return "Live"
                            Case Else : Return ""
                        End Select

                    Case "z"

                        Select Case CInt(sCat.Substring(1))
                            Case 0 : Return "Album"
                            Case 1 : Return "Liveset"
                            Case 2 : Return "Podcast"
                            Case 3 : Return "Luisterboek"
                            Case Else : Return ""
                        End Select

                    Case Else

                        Return ""

                End Select

            Case 3

                Select Case sCat.Substring(0, 1)

                    Case "a"

                        Select Case CInt(sCat.Substring(1))
                            Case 0 : Return "Windows"
                            Case 1 : Return "Macintosh"
                            Case 2 : Return "Linux"
                            Case 3 : Return "Playstation"
                            Case 4 : Return "Playstation 2"
                            Case 5 : Return "PSP"
                            Case 6 : Return "XBox"
                            Case 7 : Return "XBox 360"
                            Case 8 : Return "Gameboy Advance"
                            Case 9 : Return "Gamecube"
                            Case 10 : Return "Nintendo DS"
                            Case 11 : Return "Nintendo Wii"
                            Case 12 : Return "Playstation 3"
                            Case 13 : Return "Windows Phone"
                            Case 14 : Return "iOs"
                            Case 15 : Return "Android"
                            Case 16 : Return sIIF(bStrict, "", "Nintendo 3DS")
                            Case Else : Return ""
                        End Select

                    Case "b"

                        Select Case CInt(sCat.Substring(1))
                            Case 1 : Return "Rip"
                            Case 0 : Return sIIF(bStrict, "", "ISO")
                            Case 2 : Return "Retail"
                            Case 3 : Return "DLC"
                            Case 4 : Return "" ''"Anders"
                            Case 5 : Return "Patch"
                            Case 6 : Return "Crack"
                            Case Else : Return ""
                        End Select

                    Case "c"

                        Select Case CInt(sCat.Substring(1))
                            Case 0 : Return "Actie"
                            Case 1 : Return "Avontuur"
                            Case 2 : Return "Strategie"
                            Case 3 : Return "Rollenspel"
                            Case 4 : Return "Simulatie"
                            Case 5 : Return "Race"
                            Case 6 : Return "Vliegen"
                            Case 7 : Return "Shooter"
                            Case 8 : Return "Platform"
                            Case 9 : Return "Sport"
                            Case 10 : Return "Jeugd"
                            Case 11 : Return "Puzzel"
                            Case 12 : Return "" ''"Anders"
                            Case 13 : Return "Bordspel"
                            Case 14 : Return "Kaarten"
                            Case 15 : Return "Educatie"
                            Case 16 : Return "Muziek"
                            Case 17 : Return "Party"
                            Case Else : Return ""
                        End Select

                    Case Else

                        Return ""

                End Select

            Case 4

                Select Case sCat.Substring(0, 1)

                    Case "a"

                        Select Case CInt(sCat.Substring(1))
                            Case 0 : Return "Windows"
                            Case 1 : Return "Macintosh"
                            Case 2 : Return "Linux"
                            Case 3 : Return "OS/2"
                            Case 4 : Return "Windows Phone"
                            Case 5 : Return "Navigatie"
                            Case 6 : Return "iOs"
                            Case 7 : Return "Android"
                            Case Else : Return ""
                        End Select

                    Case "b"
                        Select Case CInt(sCat.Substring(1))
                            Case 0 : Return "Audio"
                            Case 1 : Return "Video"
                            Case 2 : Return "Grafisch"
                            Case 3 : Return sIIF(bStrict, "", "CD/DVD Tools")
                            Case 4 : Return sIIF(bStrict, "", "Media spelers")
                            Case 5 : Return sIIF(bStrict, "", "Rippers & Encoders")
                            Case 6 : Return sIIF(bStrict, "", "Plugins")
                            Case 7 : Return sIIF(bStrict, "", "Database tools")
                            Case 8 : Return sIIF(bStrict, "", "Email software")
                            Case 9 : Return "Foto"
                            Case 10 : Return sIIF(bStrict, "", "Screensavers")
                            Case 11 : Return sIIF(bStrict, "", "Skin software")
                            Case 12 : Return sIIF(bStrict, "", "Drivers")
                            Case 13 : Return sIIF(bStrict, "", "Browsers")
                            Case 14 : Return sIIF(bStrict, "", "Download managers")
                            Case 15 : Return "Download"
                            Case 16 : Return sIIF(bStrict, "", "Usenet software")
                            Case 17 : Return sIIF(bStrict, "", "RSS Readers")
                            Case 18 : Return sIIF(bStrict, "", "FTP software")
                            Case 19 : Return sIIF(bStrict, "", "Firewalls")
                            Case 20 : Return sIIF(bStrict, "", "Antivirus software")
                            Case 21 : Return sIIF(bStrict, "", "Antispyware software")
                            Case 22 : Return sIIF(bStrict, "", "Optimalisatiesoftware")
                            Case 23 : Return "Beveiliging"
                            Case 24 : Return "Systeem"
                            Case 25 : Return "" ''"Anders"
                            Case 26 : Return "Educatief"
                            Case 27 : Return "Kantoor"
                            Case 28 : Return "Internet"
                            Case 29 : Return "Communicatie"
                            Case 30 : Return "Ontwikkel"
                            Case 31 : Return "Spotnet"
                            Case Else : Return ""
                        End Select

                    Case Else

                        Return ""

                End Select

            Case Else

                Select Case sCat.Substring(0, 1)

                    Case "a"

                        Select Case CInt(sCat.Substring(1))
                            Case 0 : Return "DivX"
                            Case 1 : Return "WMV"
                            Case 2 : Return "MPG"
                            Case 3 : Return "DVD5"
                            Case 4 : Return sIIF(bStrict, "", "HD Overig")
                            Case 5 : Return "ePub"
                            Case 6 : Return "Bluray"
                            Case 7 : Return sIIF(bStrict, "", "HD-DVD")
                            Case 8 : Return sIIF(bStrict, "", "WMV HD")
                            Case 9 : Return "x264"
                            Case 10 : Return "DVD9"
                            Case Else : Return ""
                        End Select

                    Case "b"

                        Select Case CInt(sCat.Substring(1))
                            Case 4 : Return sIIF(bStrict, "", "TV")
                            Case 1 : Return sIIF(bStrict, "", "(S)VCD")
                            Case 6 : Return sIIF(bStrict, "", "Satelliet")
                            Case 2 : Return sIIF(bStrict, "", "Promo")
                            Case 3 : Return "Retail"
                            Case 7 : Return "R5"
                            Case 0 : Return "Cam"
                            Case 8 : Return sIIF(bStrict, "", "Telecine")
                            Case 9 : Return "Telesync"
                            Case 5 : Return "" '' "Anders"
                            Case 10 : Return "Scan"
                            Case Else : Return ""
                        End Select

                    Case "c"

                        Select Case CInt(sCat.Substring(1))
                            Case 0 : Return "Geen ondertitels"
                            Case 3 : Return "Engels ondertiteld (extern)"
                            Case 4 : Return sIIF(hCat <> 5, "Engels ondertiteld (ingebakken)", "Engels geschreven")
                            Case 7 : Return "Engels ondertiteld (instelbaar)"
                            Case 1 : Return "Nederlands ondertiteld (extern)"
                            Case 2 : Return sIIF(hCat <> 5, "Nederlands ondertiteld (ingebakken)", "Nederlands geschreven")
                            Case 6 : Return "Nederlands ondertiteld (instelbaar)"
                            Case 10 : Return "Engels gesproken"
                            Case 11 : Return "Nederlands gesproken"
                            Case 12 : Return sIIF(hCat <> 5, "Duits gesproken", "Duits geschreven")
                            Case 13 : Return sIIF(hCat <> 5, "Frans gesproken", "Frans geschreven")
                            Case 14 : Return sIIF(hCat <> 5, "Spaans gesproken", "Spaans geschreven")
                            Case 5 : Return ""  '' "Anders"
                            Case Else : Return ""
                        End Select

                    Case "d"

                        Select Case CInt(sCat.Substring(1))

                            Case 0 : Return "Actie"
                            Case 29 : Return "Anime"
                            Case 2 : Return "Animatie"
                            Case 28 : Return "Aziatisch"
                            Case 1 : Return "Avontuur"
                            Case 3 : Return "Cabaret"
                            Case 32 : Return "Cartoon"
                            Case 6 : Return "Documentaire"
                            Case 7 : Return "Drama"
                            Case 8 : Return "Familie"
                            Case 9 : Return "Fantasie"
                            Case 10 : Return "Filmhuis"
                            Case 12 : Return "Horror"
                            Case 33 : Return "Jeugd"
                            Case 4 : Return "Komedie"
                            Case 19 : Return "Kort"
                            Case 5 : Return "Misdaad"
                            Case 13 : Return "Muziek"
                            Case 14 : Return "Musical"
                            Case 15 : Return "Mysterie"
                            Case 21 : Return "Oorlog"
                            Case 16 : Return "Romantiek"
                            Case 17 : Return "Science Fiction"
                            Case 18 : Return "Sport"
                            Case 11 : Return sIIF(bStrict, "", "Televisie")
                            Case 20 : Return "Thriller"
                            Case 22 : Return "Western"

                            Case 23 : Return "Hetero"
                            Case 24 : Return "Homo"
                            Case 25 : Return "Lesbo"
                            Case 26 : Return "Bi"

                            Case 27 : Return "" '' "Anders"

                            Case 30 : Return "Cover"
                            Case 43 : Return "Dagblad"
                            Case 44 : Return "Tijdschrift"
                            Case 31 : Return "Stripboek"

                            Case 34 : Return "Zakelijk"
                            Case 35 : Return "Computer"
                            Case 36 : Return "Hobby"
                            Case 37 : Return "Koken"
                            Case 38 : Return "Knutselen"
                            Case 39 : Return "Handwerk"
                            Case 40 : Return "Gezondheid"
                            Case 41 : Return "Historie"
                            Case 42 : Return "Psychologie"
                            Case 45 : Return "Wetenschap"
                            Case 46 : Return "Vrouw"
                            Case 47 : Return "Religie"
                            Case 48 : Return "Roman"
                            Case 49 : Return "Biografie"
                            Case 50 : Return "Detective"
                            Case 51 : Return "Dieren"
                            Case 52 : Return ""
                            Case 53 : Return "Reizen"
                            Case 54 : Return "Waargebeurd"
                            Case 55 : Return "Non-fictie"
                            Case 57 : Return "Poezie"
                            Case 58 : Return "Sprookje"

                            Case 75 : Return sIIF(bStrict, "", "Hetero")
                            Case 74 : Return sIIF(bStrict, "", "Homo")
                            Case 73 : Return sIIF(bStrict, "", "Lesbo")
                            Case 72 : Return sIIF(bStrict, "", "Bi")

                            Case 76 : Return "Amateur"
                            Case 77 : Return "Groep"
                            Case 78 : Return "POV"
                            Case 79 : Return "Solo"
                            Case 80 : Return "Jong"
                            Case 81 : Return "Soft"
                            Case 82 : Return "Fetisj"
                            Case 83 : Return "Oud"
                            Case 84 : Return "BBW"
                            Case 85 : Return "SM"
                            Case 86 : Return "Hard"
                            Case 87 : Return "Donker"
                            Case 88 : Return "Hentai"
                            Case 89 : Return "Buiten"

                            Case Else : Return ""

                        End Select

                    Case "z"

                        Select Case CInt(sCat.Substring(1))
                            Case 0 : Return "Film"
                            Case 1 : Return "Serie"
                            Case 2 : Return "Boek"
                            Case 3 : Return "Erotiek"
                            Case Else : Return ""
                        End Select

                    Case Else

                        Return ""

                End Select

        End Select

    End Function

    Public Function TranslateCatDesc(ByVal hCat As Long, ByVal sCat As String) As String

        Select Case hCat
            Case 2
                Select Case sCat.Substring(0, 1)
                    Case "a" : Return "Formaat"
                    Case "b" : Return "Bron"
                    Case "c" : Return "Bitrate"
                    Case "d" : Return "Genre"
                    Case "z" : Return "Categorie"
                    Case Else : Return ""
                End Select
            Case 3
                Select Case sCat.Substring(0, 1)
                    Case "a" : Return "Platform"
                    Case "b" : Return "Formaat"
                    Case "c" : Return "Genre"
                    Case "z" : Return "Categorie"
                    Case Else : Return ""
                End Select
            Case 4
                Select Case sCat.Substring(0, 1)
                    Case "a" : Return "Platform"
                    Case "b" : Return "Genre"
                    Case "z" : Return "Categorie"
                    Case Else : Return ""
                End Select
            Case Else
                Select Case sCat.Substring(0, 1)
                    Case "a" : Return "Formaat"
                    Case "b" : Return "Bron"
                    Case "c" : Return "Taal"
                    Case "d" : Return "Genre"
                    Case "z" : Return "Categorie"
                    Case Else : Return ""
                End Select
        End Select

    End Function

    Public Function TranslateCatShort(ByVal hCat As Integer, ByVal sCat As Integer) As String

        Select Case hCat

            Case 2

                Select Case sCat
                    Case 0
                        Return "MP3"
                    Case 1
                        Return "WMA"
                    Case 2
                        Return "WAV"
                    Case 3
                        Return "OGG"
                    Case 4
                        Return "EAC"
                    Case 5
                        Return "DTS"
                    Case 6
                        Return "AAC"
                    Case 7
                        Return "APE"
                    Case 8
                        Return "FLAC"
                    Case Else
                        Return ""
                End Select


            Case 3

                Select Case sCat
                    Case 0
                        Return "Win"
                    Case 1
                        Return "Mac"
                    Case 2
                        Return "Linux"
                    Case 3
                        Return "PSX"
                    Case 4
                        Return "PS2"
                    Case 5
                        Return "PSP"
                    Case 6
                        Return "XBox"
                    Case 7
                        Return "360"
                    Case 8
                        Return "GBA"
                    Case 9
                        Return "GC"
                    Case 10
                        Return "NDS"
                    Case 11
                        Return "Wii"
                    Case 12
                        Return "PS3"
                    Case 13
                        Return "WP7"
                    Case 14
                        Return "iOs"
                    Case 15
                        Return "Android"
                    Case Else
                        Return ""
                End Select

            Case 4

                Select Case sCat
                    Case 0
                        Return "Win"
                    Case 1
                        Return "Mac"
                    Case 2
                        Return "Linux"
                    Case 3
                        Return "OS2"
                    Case 4
                        Return "WP7"
                    Case 5
                        Return "Navi"
                    Case 6
                        Return "iOs"
                    Case 7
                        Return "Android"
                    Case Else
                        Return ""
                End Select

            Case Else

                Select Case sCat
                    Case 0
                        Return "DivX"
                    Case 1
                        Return "WMV"
                    Case 2
                        Return "MPG"
                    Case 3
                        Return "DVD5"
                    Case 4
                        Return "HD"
                    Case 5
                        Return "ePub"
                    Case 6
                        Return "Bluray"
                    Case 7
                        Return "HD"
                    Case 8
                        Return "HD"
                    Case 9
                        Return "x264"
                    Case 10
                        Return "DVD9"
                    Case Else
                        Return ""
                End Select

        End Select

    End Function

    Public Function GetFileContents(ByVal FullPath As String, Optional ByRef ErrInfo As String = "") As String

        Dim strContents As String
        Dim objReader As StreamReader

        Try

            objReader = New StreamReader(FullPath, LatinEnc)
            strContents = objReader.ReadToEnd()
            objReader.Close()

            Return strContents

        Catch Ex As Exception

            ErrInfo = Ex.Message
            Return vbNullString

        End Try

    End Function

    Public Function LatinEnc() As Encoding

        Return Encoding.GetEncoding(&H6FAF)

    End Function

    Public Function GetLatin(ByRef zText() As Byte) As String

        Return LatinEnc.GetString(zText)

    End Function

    Public Function MakeLatin(ByVal zText As String) As Byte()

        Return LatinEnc.GetBytes(zText)

    End Function

    Public Function StripNonAlphaNumericCharacters(ByVal sText As String) As String

        If Len(sText) = 0 Then Return ""
        Return System.Text.RegularExpressions.Regex.Replace(sText, "[^A-Za-z0-9]", "").Trim

    End Function

    Private Function ConvertBytes(ByVal bytes As Long, ByVal convertTo As convTo) As Double

        If convTo.IsDefined(GetType(convTo), convertTo) Then

            Return bytes / (1024 ^ convertTo)

        Else

            Return -1 'An invalid value was passed to this function so exit

        End If

    End Function

    Public Function ConvertSize(ByVal fileSize As Long) As String

        Dim sizeOfKB As Long = 1024              ' Actual size in bytes of 1KB
        Dim sizeOfMB As Long = 1048576           ' 1MB
        Dim sizeOfGB As Long = 1073741824        ' 1GB
        Dim sizeOfTB As Long = 1099511627776     ' 1TB
        Dim sizeofPB As Long = 1125899906842624  ' 1PB

        Dim tempFileSize As Double
        Dim tempFileSizeString As String

        Dim myArr() As Char = {CChar("0"), CChar(".")}  'Characters to strip off the end of our string after formating

        If fileSize < sizeOfKB Then 'Filesize is in Bytes
            tempFileSize = ConvertBytes(fileSize, convTo.B)
            If tempFileSize = -1 Then Return Nothing 'Invalid conversion attempted so exit
            tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr) ' Strip the 0's and 1's off the end of the string
            Return Math.Round(tempFileSize) & " bytes"

        ElseIf fileSize >= sizeOfKB And fileSize < sizeOfMB Then 'Filesize is in Kilobytes
            tempFileSize = ConvertBytes(fileSize, convTo.KB)
            If tempFileSize = -1 Then Return Nothing 'Invalid conversion attempted so exit
            tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr)
            Return Math.Round(tempFileSize) & " KB"

        ElseIf fileSize >= sizeOfMB And fileSize < sizeOfGB Then ' Filesize is in Megabytes
            tempFileSize = ConvertBytes(fileSize, convTo.MB)
            If tempFileSize = -1 Then Return Nothing 'Invalid conversion attempted so exit
            tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr)
            Return Math.Round(tempFileSize, 1) & " MB"

        ElseIf fileSize >= sizeOfGB And fileSize < sizeOfTB Then 'Filesize is in Gigabytes
            tempFileSize = ConvertBytes(fileSize, convTo.GB)
            If tempFileSize = -1 Then Return Nothing
            tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr)
            Return Math.Round(tempFileSize, 1) & " GB"

        ElseIf fileSize >= sizeOfTB And fileSize < sizeofPB Then 'Filesize is in Terabytes
            tempFileSize = ConvertBytes(fileSize, convTo.TB)
            If tempFileSize = -1 Then Return Nothing
            tempFileSizeString = Format(fileSize, "Standard").TrimEnd(myArr)
            Return Math.Round(tempFileSize, 1) & " TB"

            'Anything bigger than that is silly ;)

        Else

            Return "" 'Invalid filesize so return Nothing

        End If

    End Function

    Public Function SettingsFolder() As String

        Static TheResult As String = ""
        Static DidOnce As Boolean = False

        If Not DidOnce Then

            DidOnce = True

            Dim AddFolder As String = Spotname

            If FileExists(AppPath() & "\settings.xml") Then

                TheResult = AppPath()

            Else

                Dim Myfolder As String = ""

                Myfolder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)

                If Not DirectoryExists(Myfolder & "\" & AddFolder) Then
                    If Directory.CreateDirectory(Myfolder & "\" & AddFolder) Is Nothing Then
                        Foutje("Kan '" & Myfolder & "\" & AddFolder & "' niet aanmaken!")
                    End If
                End If

                TheResult = Myfolder & "\" & AddFolder

            End If
        End If

        Return TheResult

    End Function

    Public Sub CS(Optional ByVal lMil As Integer = 1000)

        Static DidOnce As Boolean = False

        If Not DidOnce Then
            DidOnce = True
            Try
                My.Application.CloseSplash(lMil)
            Catch ex As Exception
                Foutje("CS: " & ex.Message)
                Foutje("CS: " & ex.Message)
            End Try
        End If

    End Sub

    Public Sub Foutje(ByVal sMsg As String, Optional ByVal sCaption As String = "Oeps")

        CS(0)
        Debug.Print(sMsg)

        Try
            MsgBox(sMsg, CType(MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, MsgBoxStyle), sCaption)
        Catch
        End Try

    End Sub

    Public Function GetFolder(ByVal pGuid As Guid) As String

        Dim txtOldTmp As String = ""
        Dim ptPath As IntPtr

        If SHGetKnownFolderPath(pGuid, 0, IntPtr.Zero, ptPath) = 0 Then
            txtOldTmp = System.Runtime.InteropServices.Marshal.PtrToStringUni(ptPath)
            System.Runtime.InteropServices.Marshal.FreeCoTaskMem(ptPath)
        End If

        GetFolder = txtOldTmp

    End Function

    Public Function AppPath() As String

        Static TheResult As String = ""
        Static DidOnce As Boolean = False

        If Not DidOnce Then

            Try

                TheResult = System.Windows.Forms.Application.StartupPath
                DidOnce = True

            Catch ex As Exception
                Foutje("AppPath")
            End Try

        End If

        Return TheResult

    End Function

    Public Function DownDir() As String

        If Len(Trim(My.Settings.DownloadFolder)) > 0 Then
            If DirectoryExists(My.Settings.DownloadFolder) Then
                Return My.Settings.DownloadFolder
            End If
        End If

        If (Environment.OSVersion.Version.Major >= 6) Then
            Return GetFolder(New Guid("374DE290-123F-4565-9164-39C4925E467B"))
        Else
            Return Environment.ExpandEnvironmentVariables("%USERPROFILE%") & "\Downloads"
        End If

    End Function

    Public Function MakeFilename(ByVal sIn As String) As String

        Return sIn.Replace(Chr(34), "").Replace("*", "").Replace(":", "").Replace("<", "").Replace(">", "").Replace("?", "").Replace("\", "").Replace("/", "").Replace("|", "").Replace(".", "").Replace("%", "").Replace("[", "").Replace("]", "").Replace(";", "").Replace("=", "").Replace(",", "")

    End Function

    Public Sub LaunchBrowser(ByVal sUrl As String)

        Try
            System.Diagnostics.Process.Start(sUrl)
        Catch
            Foutje("LaunchBrowser:: " & Err.Description)
        End Try

    End Sub

    Private Function CreateHash(ByVal sLeft As String, ByVal sRight As String) As String

        Dim RetBytes() As Byte
        Dim CharCnt As Integer = 62
        Dim lByteCount As Integer = 5

        Dim SHA As SHA1CryptoServiceProvider = New SHA1CryptoServiceProvider()

        Dim OrgBytes() As Byte = MakeLatin(sLeft)
        Dim OrgBytes2() As Byte = MakeLatin(sRight)

        Dim OrgLen As Integer = UBound(OrgBytes) + 1
        Dim OrgLen2 As Integer = UBound(OrgBytes2) + 1

        Dim TryBytes(OrgLen + OrgLen2 - 2 + lByteCount) As Byte

        For Zl As Integer = 0 To UBound(OrgBytes)
            TryBytes(Zl) = OrgBytes(Zl)
        Next

        For Zl As Integer = UBound(TryBytes) - (OrgLen2 - 1) To UBound(TryBytes)
            TryBytes(Zl) = OrgBytes2(Zl - (UBound(TryBytes) - (OrgLen2 - 1)))
        Next

        Dim xPos1 As Integer = UBound(TryBytes) - 3 - OrgLen2
        Dim xPos2 As Integer = UBound(TryBytes) - 2 - OrgLen2
        Dim xPos3 As Integer = UBound(TryBytes) - 1 - OrgLen2
        Dim xPos4 As Integer = UBound(TryBytes) - OrgLen2

        Dim zR As New Random

        Dim Xp1B(CharCnt - 1) As Byte
        Dim Xp2B(CharCnt - 1) As Byte
        Dim Xp3B(CharCnt - 1) As Byte
        Dim Xp4B(CharCnt - 1) As Byte

        For i0 As Integer = 0 To CharCnt - 1
            Xp1B(i0) = CByte(i0)
            Xp2B(i0) = CByte(i0)
            Xp3B(i0) = CByte(i0)
            Xp4B(i0) = CByte(i0)
        Next

        Dim r As Integer = 0
        Dim t As Byte = 0

        For i0 As Integer = 0 To CharCnt - 1

            r = zR.Next(0, CharCnt)

            t = Xp1B(i0)
            Xp1B(i0) = Xp1B(r)
            Xp1B(r) = t

            r = zR.Next(0, CharCnt)

            t = Xp2B(i0)
            Xp2B(i0) = Xp2B(r)
            Xp2B(r) = t

            r = zR.Next(0, CharCnt)

            t = Xp3B(i0)
            Xp3B(i0) = Xp3B(r)
            Xp3B(r) = t

            r = zR.Next(0, CharCnt)

            t = Xp4B(i0)
            Xp4B(i0) = Xp4B(r)
            Xp4B(r) = t
        Next

        For i0 As Integer = 0 To CharCnt - 1
            Xp1B(i0) = GetBaseChar(Xp1B(i0))
            Xp2B(i0) = GetBaseChar(Xp2B(i0))
            Xp3B(i0) = GetBaseChar(Xp3B(i0))
            Xp4B(i0) = GetBaseChar(Xp4B(i0))
        Next

        For i1 As Integer = 0 To CharCnt - 1
            For i2 As Integer = 0 To CharCnt - 1
                For i3 As Integer = 0 To CharCnt - 1
                    For i4 As Integer = 0 To CharCnt - 1

                        TryBytes(xPos1) = Xp1B(i1)
                        TryBytes(xPos2) = Xp2B(i2)
                        TryBytes(xPos3) = Xp3B(i3)
                        TryBytes(xPos4) = Xp4B(i4)

                        RetBytes = SHA.ComputeHash(TryBytes)

                        If (RetBytes(0) = 0) Then

                            If (RetBytes(1) = 0) Then

                                Return GetLatin(TryBytes)

                            End If

                        End If

                    Next
                Next
            Next
        Next

        Throw New Exception("Error 422")
        Return Nothing

    End Function

    Private Function GetBaseChar(ByVal lIndex As Byte) As Byte

        Select Case lIndex
            Case 0 To 25
                GetBaseChar = CByte(65 + lIndex)
            Case 26 To 51
                GetBaseChar = CByte(97 + (lIndex - 26))
            Case 52 To 62
                GetBaseChar = CByte(48 + (lIndex - 52))
            Case Else
                GetBaseChar = CByte(65)
        End Select

    End Function

    Friend Function sIIF(ByVal bExpression As Boolean, ByVal sTrue As String, ByVal sFalse As String) As String

        If bExpression Then
            Return sTrue
        Else
            Return sFalse
        End If

    End Function

    Friend Function GetKey() As RSACryptoServiceProvider

        Static cKey As RSACryptoServiceProvider

        If cKey Is Nothing Then

            Dim CP As New CspParameters

            CP.KeyContainerName = Spotname & " User Key"
            CP.Flags = CType(CspProviderFlags.NoPrompt + CspProviderFlags.UseArchivableKey, CspProviderFlags)

            cKey = New RSACryptoServiceProvider(384, CP)

        End If

        Return cKey

    End Function

    Friend Function GetAvatar() As Byte()

        If Len(My.Settings.Avatar) = 0 Then Return Nothing

        Try

            Dim oMS As New MemoryStream()

            Dim zIn() As Byte = Convert.FromBase64String(My.Settings.Avatar)
            oMS.Write(zIn, 0, zIn.Length)

            Dim PL2 As BitmapFrame = BitmapFrame.Create(oMS, BitmapCreateOptions.DelayCreation, BitmapCacheOption.None)
            If Not (PL2.PixelWidth > 1 And PL2.PixelHeight > 1) Then Return Nothing

            Return zIn

        Catch
        End Try

        Return Nothing

    End Function

    Friend Function GetModulus() As String

        Return Convert.ToBase64String(GetKey.ExportParameters(False).Modulus)

    End Function

    Public Function CreateKeys() As Boolean

        Dim sKey(10) As String

        Try

            If Not System.IO.File.Exists(SettingsFolder() & "\keys.xml") Then

                sKey(2) = "ys8WSlqonQMWT8ubG0tAA2Q07P36E+CJmb875wSR1XH7IFhEi0CCwlUzNqBFhC+P"
                sKey(3) = "uiyChPV23eguLAJNttC/o0nAsxXgdjtvUvidV2JL+hjNzc4Tc/PPo2JdYvsqUsat"
                sKey(4) = "1k6RNDVD6yBYWR6kHmwzmSud7JkNV4SMigBrs+jFgOK5Ldzwl17mKXJhl+su/GR9"

                Dim SK As New StreamWriter(SettingsFolder() & "\keys.xml", False, System.Text.Encoding.UTF8)

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

            Foutje("Create_Keys: " & ex.Message)
            Return False

        End Try

        Return True

    End Function

    Public Function CreateList(ByVal sFile As String, ByVal cList As List(Of ListItem)) As Boolean

        Try

            Dim SK As New StreamWriter(SettingsFolder() & "\" & sFile, False, System.Text.Encoding.UTF8)

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

            Foutje("Create_List: " & ex.Message)

        End Try

        Return False

    End Function

    Private Function DownloadString(ByVal zUrl As String, ByRef zError As String) As String

        Dim oWeb As System.Net.WebClient

        Try

            oWeb = New System.Net.WebClient
            Return (oWeb.DownloadString(zUrl))

        Catch ex As Exception
            zError = ex.Message
            Return vbNullString
        End Try

    End Function

    Private Function UpdateList(ByVal zUrl As String, ByVal sFile As String) As Boolean

        Try

            Dim sXml As String = DownloadString(zUrl, "")

            If sXml Is Nothing Then Return False
            If Len(sXml) = 0 Then Return False
            If Not ValidList(sXml) Then Return False

            Dim SK As New StreamWriter(SettingsFolder() & "\" & sFile, False, System.Text.Encoding.UTF8)
            SK.Write(sXml)
            SK.Close()

            Return True

        Catch
        End Try

        Return False

    End Function

    Friend Function ValidList(ByVal sXML As String) As Boolean

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

    Private Function LoadList(ByVal sFile As String, ByVal ToList As HashSet(Of String)) As Boolean

        Try

            Dim root As XmlElement
            Dim doc As New XmlDocument

            doc.XmlResolver = Nothing
            doc.Load(SettingsFolder() & "\" & sFile)
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

    Friend Function GetIcon(ByVal sKey As String) As Image

        Try
            Dim icon As Image = New Image
            icon.Source = New BitmapImage(New Uri("pack://application:,,,/Spotnet;component/Images/" & sKey & ".ico"))
            Return icon
        Catch
            Return Nothing
        End Try

    End Function

    Friend Function GetImage(ByVal sKey As String) As ImageSource

        Try
            Return New BitmapImage(New Uri("pack://application:,,,/Spotnet;component/Images/" & sKey))
        Catch
            Return Nothing
        End Try

    End Function

    Public Function GetColorFromHex(ByVal myColor As String) As SolidColorBrush

        Try

            Return New SolidColorBrush(
                Color.FromArgb(
                    Convert.ToByte(myColor.Substring(1, 2), 16),
                    Convert.ToByte(myColor.Substring(3, 2), 16),
                    Convert.ToByte(myColor.Substring(5, 2), 16),
                    Convert.ToByte(myColor.Substring(7, 2), 16)))

        Catch ex As Exception

            Return New SolidColorBrush(Colors.White)

        End Try

    End Function

    Public Function ReturnInfo(ByVal ExtCat As Integer) As String

        Dim FindCat As Char
        Dim sCat As String = CStr(CInt(ExtCat))

        Dim lCat As Byte = CByte(sCat.Substring(0, 1))
        Dim zCat As Byte = CByte(sCat.Substring(1))

        If zCat > 98 Then Return ""

        Select Case lCat
            Case 3
                FindCat = "c"c
            Case 4
                FindCat = "b"c
            Case Else
                FindCat = "d"c
        End Select

        Return TranslateCat(lCat, FindCat & zCat)

    End Function

    Public Function TranslateInfo(ByVal hCat As Integer, ByVal sCats As String) As Byte

        Dim FindCat As Char
        Dim xInd As Byte = 0
        Dim SubCat As Long = 0
        Dim lPos As Integer = -1
        Dim sFind As String = ""
        Dim sRet As String = ""
        Dim MaxCat As Byte = 100

        Static bCat0(MaxCat) As Boolean
        Static bCat1(MaxCat) As Boolean
        Static bCat2(MaxCat) As Boolean
        Static bCat3(MaxCat) As Boolean

        Static DidOnce As Boolean = False

        If Not DidOnce Then

            For i As Byte = 0 To MaxCat
                sRet = TranslateCat(1, "d" & i)
                bCat0(i) = (Len(sRet) > 0)
            Next

            For i As Byte = 0 To MaxCat
                sRet = TranslateCat(2, "d" & i)
                bCat1(i) = (Len(sRet) > 0)
            Next

            For i As Byte = 0 To MaxCat
                sRet = TranslateCat(3, "c" & i)
                bCat2(i) = (Len(sRet) > 0)
            Next

            For i As Byte = 0 To MaxCat
                sRet = TranslateCat(4, "b" & i)
                bCat3(i) = (Len(sRet) > 0)
            Next

            DidOnce = True

        End If

        Try

            Select Case hCat
                Case 3
                    FindCat = "c"c
                Case 4
                    FindCat = "b"c
                Case Else
                    FindCat = "d"c
            End Select

            Do

                lPos = sCats.IndexOf(FindCat, lPos + 1)

                If lPos = -1 Then Exit Do

                xInd = CByte(Val(sCats.Substring(lPos + 1, 2)))

                If xInd > MaxCat Then Continue Do

                If hCat = 6 And xInd = 11 Then Continue Do

                If hCat = 9 Then

                    Select Case xInd

                        Case 23, 24, 25, 26, 72, 73, 74, 75

                            If sCats.IndexOf(FindCat, lPos + 1) > -1 Then Continue Do

                    End Select

                End If

                Select Case hCat

                    Case 1
                        If Not bCat1(xInd) Then Continue Do

                    Case 2
                        If Not bCat2(xInd) Then Continue Do

                    Case 3
                        If Not bCat3(xInd) Then Continue Do

                    Case Else
                        If Not bCat0(xInd) Then Continue Do

                End Select

                Return xInd

            Loop

        Catch ex As Exception
        End Try

        Return 99

    End Function

    Friend Function CreateHeader(ByVal sTitle As String, ByVal sIcon As String, Optional ByVal ISX As ImageSource = Nothing) As StackPanel

        Dim KL As New StackPanel
        KL.Orientation = Orientation.Horizontal

        Dim kli As New Image
        If ISX Is Nothing Then kli.Source = GetImage(sIcon) Else kli.Source = ISX
        kli.VerticalAlignment = Windows.VerticalAlignment.Bottom

        Dim klt As New TextBlock
        klt.Text = sTitle.Trim & " "
        klt.Margin = New Thickness(4, 0, 0, 0)
        klt.VerticalAlignment = Windows.VerticalAlignment.Center

        KL.Children.Add(kli)
        KL.Children.Add(klt)

        Return KL

    End Function

    Friend Function GetHeader(ByVal Zt As Object) As String

        Dim kd As StackPanel = CType(Zt, StackPanel)
        Dim pd As TextBlock = CType(kd.Children(1), TextBlock)
        Return pd.Text

    End Function

    Friend Function GetHeaderIcon(ByVal Zt As Object) As ImageSource

        Dim kd As StackPanel = CType(Zt, StackPanel)
        Dim pd As Image = CType(kd.Children(0), Image)
        Return pd.Source

    End Function

    Public Function FindNZB(ByVal sName As String, ByVal sNewsGroup As String, ByVal sTitle As String) As String

        Dim ssNzb As String = ""
        ''ssNzb = Spots.FindNZB(sName, 600, sNewsGroup, 999, True, False)

        If (Len(ssNzb) = 0) Then

            MsgBox("Kan NZB niet vinden, klik de bestandsnaam om handmatig te zoeken.", CType(MsgBoxStyle.Information + MsgBoxStyle.OkOnly, MsgBoxStyle), "NZB niet gevonden!")
            Return vbNullString

        Else

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
                objReader.Write(ssNzb)
                objReader.Close()
            Catch Ex As Exception
                Foutje("Fout tijdens het schrijven.")
                Return vbNullString
            End Try

            Return TmpFile

        End If

        Return vbNullString

    End Function

    Friend Function CreateMsgID(Optional ByVal sPrefix As String = "") As String

        Dim ZL(7) As Byte
        Dim ZK As New Random
        Dim sDomain As String = "spot.net"

        ZK.NextBytes(ZL)

        Dim Span As TimeSpan = (DateTime.UtcNow - EPOCH)
        Dim CurDate As Integer = CInt(Span.TotalSeconds)

        Dim sRandom As String = Convert.ToBase64String(ZL) & Convert.ToBase64String(BitConverter.GetBytes(CurDate))
        sRandom = sRandom.Replace("/", "s").Replace("+", "p").Replace("=", "")

        If Len(sPrefix) = 0 Then
            Return CreateHash("<" & sRandom, "@" & sDomain & ">")
        Else
            Return CreateHash("<" & sPrefix.Replace(".", "") & ".0." & sRandom & ".", "@" & sDomain & ">")
        End If

    End Function

    Friend Function GetCrc() As CRC32

        Static CC As CRC32 = Nothing

        If CC Is Nothing Then CC = New CRC32

        Return CC

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

    Public ReadOnly Property LastPosition(ByVal db As SqlDB, ByVal sTable As String) As Long

        Get

            Dim sErr As String = ""

            Try

                Dim dbCmd As DbCommand = db.CreateCommand
                dbCmd.CommandText = "SELECT MAX(rowid) FROM " & sTable

                Dim tObj As Object = dbCmd.ExecuteScalar()

                If IsDBNull(tObj) Then Return -1

                Return CType(tObj, Long)

            Catch ex As Exception

                Throw New Exception("LastPosition: " & ex.Message)

            End Try

        End Get

    End Property

    Friend Function BlackList() As HashSet(Of String)

        Dim sLocal As String = "blacklist.xml"

        If iBlackList Is Nothing Then

            If Not System.IO.File.Exists(SettingsFolder() & "\" & sLocal) Then

                CreateList(sLocal, New List(Of ListItem))

            End If

            iBlackList = New HashSet(Of String)

            LoadList(sLocal, iBlackList)

            UpdateBlacklist()

        End If

        Return iBlackList

    End Function

    Friend Sub UpdateBlacklist()

        Dim sServer As String = "blacklist.server.xml"

        If Len(My.Settings.BlacklistURL) > 0 Then

            UpdateList(AddHttp(My.Settings.BlacklistURL), sServer)

            If System.IO.File.Exists(SettingsFolder() & "\" & sServer) Then

                LoadList(sServer, iBlackList)

            End If

        End If

    End Sub

    Private Function AddToList(ByVal sName As String, ByVal sKey As String, ByVal sFile As String) As Boolean

        Try

            Dim root As XmlElement
            Dim doc As New XmlDocument

            doc.XmlResolver = Nothing
            doc.Load(SettingsFolder() & "\" & sFile)
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

            If FileExists(SettingsFolder() & "\" & sFile) Then
                System.IO.File.SetAttributes(SettingsFolder() & "\" & sFile, IO.FileAttributes.Normal)
            End If

            doc.Save(SettingsFolder() & "\" & sFile)

            Return True

        Catch
        End Try

        Return False

    End Function

    Private Function RemoveFromList(ByVal sKey As String, ByVal sFile As String) As Boolean

        Try

            Dim root As XmlElement
            Dim doc As New XmlDocument
            Dim bDidSome As Boolean = False

            If Not FileExists(SettingsFolder() & "\" & sFile) Then Return False

            doc.XmlResolver = Nothing
            doc.Load(SettingsFolder() & "\" & sFile)
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

            If FileExists(SettingsFolder() & "\" & sFile) Then
                System.IO.File.SetAttributes(SettingsFolder() & "\" & sFile, IO.FileAttributes.Normal)
            End If

            doc.Save(SettingsFolder() & "\" & sFile)

            Return True

        Catch
        End Try

        Return False

    End Function

    Friend Function AddWhite(ByVal sName As String, ByVal sKey As String) As Boolean

        iWhiteList.Add(sKey)
        Return AddToList(sName, sKey, "whitelist.xml")

    End Function

    Friend Function AddBlack(ByVal sName As String, ByVal sKey As String) As Boolean

        iBlackList.Add(sKey)
        Return AddToList(sName, sKey, "blacklist.xml")

    End Function

    Friend Function RemoveWhite(ByVal sKey As String) As Boolean

        iWhiteList.Remove(sKey)

        RemoveFromList(sKey, "whitelist.server.xml")
        Return RemoveFromList(sKey, "whitelist.xml")

    End Function

    Friend Function RemoveBlack(ByVal sKey As String) As Boolean

        iBlackList.Remove(sKey)

        RemoveFromList(sKey, "blacklist.server.xml")
        Return RemoveFromList(sKey, "blacklist.xml")

    End Function

    Friend Function WhiteList() As HashSet(Of String)

        Dim sLocal As String = "whitelist.xml"

        If iWhiteList Is Nothing Then

            If Not System.IO.File.Exists(SettingsFolder() & "\" & sLocal) Then

                Dim dList As New List(Of ListItem)

                dList.Add(New ListItem("1wt6jlePL/IADm4wL8lMqHaGVznPTiUvcovAtj3eCgvt3wTyM9Fd8ptx8+xzmAHL", "Albertina"))
                dList.Add(New ListItem("ynakBYOJnwLBuXQZvglD1N/uZ0mZqYad9dKX9KxyOe2mPoEZIE8Y/x93U8VL4tnv", "Bacoben1"))
                dList.Add(New ListItem("6ibY+eDYDwXOjV992fdCqhE0V0B2rRwqvxmoodPlpgjSshPCUgVjTHqpoC1AzbqR", "Boaz"))
                dList.Add(New ListItem("vaaHp9taPnRVbYZaa5etSK6y4Caft5aOrnzqfjPljgD2UE/89TBz6JbA/NeJpK+p", "BOB1961"))
                dList.Add(New ListItem("s7xw10e0wq6dZgrkD59T9F/lj0zSaht0Zv0gYVvS2gR7I4VPjo/TrqxhwSP3by//", "Biky"))
                dList.Add(New ListItem("zNOkGYubV87uJaL1KIqqHHs+nKWNwhD0yNEu0Mz4TKBVkDkxdTB8RvcAa79tMyaL", "Blowan"))
                dList.Add(New ListItem("0pGKk73HQkkj1waqHSjuMtpqAuAhItXNYXOQXHQL+rqORxzqMMoQeg523iJKUbvf", "Bradje"))
                dList.Add(New ListItem("snmypn4sZq+N4tn+UT6IFPn9Ii67iteD/T/weYVVQbQWvui4M1SSUxaqvIFQtQ8l", "CaptainSalvo"))
                dList.Add(New ListItem("58UoKbJ7JgNbRFJJqpdwO3MYKexHlkkUt6KfZvP7lykUNHRm/sZssM4o2jUm6TUh", "CaptainSalvo"))
                dList.Add(New ListItem("rCpFxtuo9ijYWTg4WpDnQQO2dVQGSlhGamuUmWCrpilfEbWKNLap+EFnNEqCHdbF", "Dick42"))
                dList.Add(New ListItem("z+U5teGdCtU0MVePPPZu1APEJfpSAPNh/RR1EyBXRD1G8d73M+qJZJqfJUL9smUF", "Falang01"))
                dList.Add(New ListItem("ru3rhWGBsx4dCglEwjE3bL9nVfH2gJVS0kb0OrXQTceeMXLDVb4rsuA+ty85M3If", "Hagenees1978"))
                dList.Add(New ListItem("xC2V+4i7J07fm6+ND+Mr5hvD359l2R/bkeOt2cGUpeFznxhItdMEVJKDNthKFNIb", "Hagenees1978"))
                dList.Add(New ListItem("ySv0wJaY8WQPb1KUkJeOVr4dGqR2UoxaOxsnqYmcgkbiPhigkb235eVvoIj4AVM7", "Hannes3"))
                dList.Add(New ListItem("z/4mkqzLE27ur8iNOTerBbFK37//itkNa5APDIRLTQ3gBJZORgOcqT+51lw2qnQx", "HendrikjeStoffel"))
                dList.Add(New ListItem("twJLKIJYDQvTGhk3hnLSWdgE9oXkH/RypTAI7Bo2rBHkH5FfL/FOJEvOp/MVRWFP", "Inge2222"))
                dList.Add(New ListItem("reZxfDPBE/Bxqa63PW4LFiDTh6xl7w1Sh3eoYmUbYbiI8AbmtWNmWAWjC6mHef+b", "Kaj7"))
                dList.Add(New ListItem("szsAIT5lVEzonnwg81DoU/44KTXkdIYrAdAFpoB/99Fw0VC6QVad7PRKgDPFeDW5", "kww"))
                dList.Add(New ListItem("qIxm7gFn8z6eIheHbstSa0vEhciwEMzNMjYlvBXJEBmivtcfrTXXz57VMfIDtKZB", "kww"))
                dList.Add(New ListItem("4ci3BuoC+JHlHVTxYacoEmk7rXGnrRlmgp1zuNO/wrtX0M0ixhK1MUlMMZIaVJ39", "Oldtimer"))
                dList.Add(New ListItem("rla55FY/Gm1DgPFwo4+HgMq8bElbjW9W8dIBUFun3ujfujp89p07LAQkS32FWQNb", "Ricardoo"))
                dList.Add(New ListItem("qp1ja8wjPDlh7aEssytHTflMCeKLF1TDoZlA41Qp9rkifx+qz9oY21FqZxgOiQgN", "Ricardoo"))
                dList.Add(New ListItem("+QIm6ZjUIY8Jgn0venbvGoik2hZyPZpNlXJrGlCbQgRndiN4apVb9awMsp2YGY5j", "SubmarinesSpot"))
                dList.Add(New ListItem("p2T1o0E6djKXSBqv8sRPLVsKxZnZOzuQgzY25QBdF5l5+5El81ziGD+5RBuUXSkT", "SubmarinesSpot"))
                dList.Add(New ListItem("0mqEBWp/z9l8W15lwntuXNxpcY04o96/MxGe4OCg6dFCjzQ6g8kSej/QoL9tkhB/", "Sophia1949"))
                dList.Add(New ListItem("9lfEiCusAUMTMqCOi6sc6P2IYoDslFbGEIYZ16ku6Nqrclc5oyLE7wz2fUI0RJvx", "Trein1600"))
                dList.Add(New ListItem("yf0oZC/mJLo0iHunzKn1YyvPCyI6r/ACTNAG3K53BzF4efYWe37EC9P4nRmHEDPJ", "xxxwebwatchers"))
                dList.Add(New ListItem("zxCiZ9F9yZ7DdEPj+1Ta/nQl679amRgc+BcmFuRpWvt9VjnHzY7dUTMPUavB8jUN", "Y0os"))
                dList.Add(New ListItem("u9bdM+NQl4OPvhi4GHiRvyvDRuTVBemAeAh70lpIWGRqiv03hDvI7W53FuQk3rDX", "Zoutoplossing"))
                dList.Add(New ListItem("uH19iDBeTjye6rhOi4uLR+T59MThUBQNL0ZgQhsX6BQqQxZNYflwccud9ZN64Rb5", "Zoutoplossing"))

                CreateList(sLocal, dList)

            End If

            iWhiteList = New HashSet(Of String)

            LoadList(sLocal, iWhiteList)

            UpdateWhitelist()

        End If

        Return iWhiteList

    End Function

    Friend Sub UpdateWhitelist()

        Dim sServer As String = "whitelist.server.xml"

        If Len(My.Settings.WhitelistURL) > 0 Then

            UpdateList(AddHttp(My.Settings.WhitelistURL), sServer)

            If System.IO.File.Exists(SettingsFolder() & "\" & sServer) Then

                LoadList(sServer, iWhiteList)

            End If

        End If

    End Sub

    Friend Function LoadKeys() As String()

        Static sKey(10) As String
        Static KeysLoaded As Boolean = False

        If KeysLoaded Then Return sKey

        Try

            Dim sLocal As String = "keys.xml"

            If Len(My.Settings.KeysURL) > 0 Then

                UpdateList(AddHttp(My.Settings.KeysURL), sLocal)

            End If

            If Not System.IO.File.Exists(SettingsFolder() & "\" & sLocal) Then

                CreateKeys()

            End If

            Dim root As XmlElement
            Dim doc As New XmlDocument
            Dim child2 As XmlElement

            doc.XmlResolver = Nothing
            doc.Load(SettingsFolder() & "\" & sLocal)
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

    Friend Function MakeUnique(ByVal sModulus As String) As String

        If Len(sModulus) = 0 Then Return "Onbekend"

        Try
            Return StripNonAlphaNumericCharacters(Convert.ToBase64String(BitConverter.GetBytes(GetCrc.Calculate(Convert.FromBase64String(sModulus)))))
        Catch
            Return "Onbekend"
        End Try

    End Function

    Friend Function MakeMD5(ByVal sModulus As String) As String

        Static xMD5 As New MD5CryptoServiceProvider()

        If Len(sModulus) = 0 Then Return "Onbekend"

        Try

            Dim bBytes() As Byte = xMD5.ComputeHash(Convert.FromBase64String(sModulus))
            Dim sBuilder As New StringBuilder((UBound(bBytes) + 1) * 2, (UBound(bBytes) + 1) * 2)

            For i = 0 To bBytes.Length - 1
                sBuilder.Append(bBytes(i).ToString("x2"))
            Next

            Return sBuilder.ToString()

        Catch
            Return "Onbekend"
        End Try

    End Function

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

    Friend Function SpotTheme() As String

        Dim sFile As String = "spot"
        Dim objReader As StreamWriter

        sFile = SettingsFolder() & "\Theme\" & sFile & ".htm"

        If Not FileExists(sFile) Then
            Try

                If Not DirectoryExists(SettingsFolder() & "\Theme") Then
                    Directory.CreateDirectory(SettingsFolder() & "\Theme")
                End If

                objReader = New StreamWriter(sFile, False, LatinEnc)
                objReader.Write(My.Resources.spot.ToString())
                objReader.Close()

            Catch Ex As Exception
            End Try
        End If

        Return sFile

    End Function

    Friend Function CommentTheme() As String

        Dim sFile As String = "comment"
        Dim objReader As StreamWriter

        sFile = SettingsFolder() & "\Theme\" & sFile & ".htm"

        If Not FileExists(sFile) Then

            Try

                If Not DirectoryExists(SettingsFolder() & "\Theme") Then
                    Directory.CreateDirectory(SettingsFolder() & "\Theme")
                End If

                objReader = New StreamWriter(sFile, False, LatinEnc)
                objReader.Write(My.Resources.comment.ToString())
                objReader.Close()

            Catch Ex As Exception
            End Try
        End If

        Return sFile

    End Function

    Friend Sub ShowOnce(ByVal sMsg As String, ByVal sTitle As String)

        Static DidList As New HashSet(Of String)

        If DidList.Add(sMsg) Then
            MsgBox(sMsg, MsgBoxStyle.Information, sTitle)
        End If

    End Sub

    Friend Function IsSearchQuery(ByVal sQuery As String) As Boolean

        Dim sTest As String = sQuery.ToLower.Trim
        Return sTest.Contains(" match ")

    End Function

End Module