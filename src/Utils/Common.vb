Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.Threading
Imports System.Security.Cryptography
Imports System.Windows.Media.Imaging

Imports Spotlib
Imports Spotbase.Spotbase

Friend Class SpotInfo

    Public Spot As SpotEx
    Public TabLoaded As Boolean

End Class

Friend Class UrlInfo

    Public URL As String
    Public Title As String
    Public TabLoaded As Boolean

End Class

Public Enum ServerType
    Headers
    Upload
    Download
End Enum

Public Class ServerInfo

    Public Port As Integer = 119
    Public SSL As Boolean = False
    Public Username As String = ""
    Public Password As String = ""
    Public Server As String = ""
    Public Connections As Integer = 1

End Class

Public Class ListItem

    Public Key As String
    Public Name As String

    Public Sub New(ByVal sKey As String, ByVal sName As String)

        Key = sKey
        Name = sName

    End Sub

End Class

Public Class ProviderItem

    Public Name As String = ""
    Public Address As String = ""
    Public Port As Long

    Public Overrides Function ToString() As String
        Return Name
    End Function

End Class

Friend Class SabSlots

    Public DataSlots As List(Of String())
    Public HistorySlots As List(Of String())

End Class
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

    'Public ReadOnly Property LastPosition(ByVal db As SqlDB, ByVal sTable As String) As Long

    '    Get

    '        Dim sErr As String = ""

    '        Try

    '            Dim dbCmd As DbCommand = db.CreateCommand
    '            dbCmd.CommandText = "SELECT MAX(rowid) FROM " & sTable

    '            Dim tObj As Object = dbCmd.ExecuteScalar()

    '            If IsDBNull(tObj) Then Return -1

    '            Return CType(tObj, Long)

    '        Catch ex As Exception

    '            Throw New Exception("LastPosition: " & ex.Message)

    '        End Try

    '    End Get

    'End Property

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