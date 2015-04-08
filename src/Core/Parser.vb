Imports System.IO
Imports System.Text
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions

Imports Spotlib
Imports Spotbase.Spotbase

Friend Class SpotParser

    Friend Shared Function ParseSpot(ByVal xSpot As SpotEx, ByVal FontSize As Byte) As String

        Static SpotHTM As String = Nothing

        If SpotHTM Is Nothing Then

            Dim sR As New StreamReader(SpotTheme)
            SpotHTM = sR.ReadToEnd()

        End If

        Return SpotHTM.Replace("[SN:NICK]", StripNonAlphaNumericCharacters(My.Settings.Nickname)).Replace("[SN:AVATAR]", GetAvatar(xSpot.User.Avatar, xSpot.User.Modulus)).Replace("[SN:FONTSIZE]", CStr(FontSize)).Replace("[SN:FILESIZE]", CStr(xSpot.Filesize)).Replace("[SN:STAMP]", CStr(xSpot.Stamp)).Replace("[SN:MODULUS]", xSpot.User.Modulus).Replace("[SN:CAT]", HtmlEncode(Utils.CatDesc(xSpot.Category))).Replace("[SN:PATH]", SettingsFolder.Replace(Chr(34), Chr(34) & Chr(34))).Replace("[SN:TAG]", HtmlEncode(StripNonAlphaNumericCharacters(xSpot.Tag))).Replace("[SN:FROM]", HtmlEncode(StripNonAlphaNumericCharacters(xSpot.Poster))).Replace("[SN:UNIQUE]", sIIF(xSpot.User.ValidSignature, MakeUnique(xSpot.User.Modulus), "Onbekend")).Replace("[SN:WHITELIST]", sIIF(xSpot.User.ValidSignature And WhiteList.Contains(xSpot.User.Modulus), "True", "False")).Replace("[SN:WEB]", SafeHref(xSpot.Web).Replace("[SN:", "[SN:]")).Replace("[SN:SUBCATS]", HtmlEncode(xSpot.SubCats).Replace("[SN:", "[SN:]")).Replace("[SN:TITLE]", HtmlEncode(xSpot.Title).Replace("[SN:", "[SN:]")).Replace("[SN:STATS]", "").Replace("[SN:IMGX]", CStr(xSpot.ImageWidth)).Replace("[SN:IMGY]", CStr(xSpot.ImageHeight)).Replace("[SN:INFO]", MakeTable(xSpot)).Replace("[SN:DESC]", MakeDesc(xSpot))

    End Function

    Friend Shared Function ParseComment(ByVal xComment As Comment, ByVal sClass As String, ByVal sTooltip As String) As String

        Static CommentHTM As String = Nothing

        If CommentHTM Is Nothing Then

            Dim sR As New StreamReader(CommentTheme)
            CommentHTM = sR.ReadToEnd()

        End If

        Return CommentHTM.Replace("[SN:AVATAR]", GetAvatar(xComment.User.Avatar, xComment.User.Modulus)).Replace("[SN:CLASS]", sClass).Replace("[SN:TOOLTIP]", sTooltip).Replace("[SN:STAMP]", CStr(CType(xComment.Created - EPOCH, TimeSpan).TotalSeconds)).Replace("[SN:DATE]", xComment.Created.ToLocalTime.ToString("%d MMM | HH:mm")).Replace("[SN:MODULUS]", xComment.User.Modulus).Replace("[SN:PATH]", SettingsFolder.Replace(Chr(34), Chr(34) & Chr(34))).Replace("[SN:FROM]", HtmlEncode(xComment.From)).Replace("[SN:UNIQUE]", MakeUnique(xComment.User.Modulus)).Replace("[SN:WHITELIST]", sIIF(WhiteList.Contains(xComment.User.Modulus), "True", "False")).Replace("[SN:DESC]", MakeBody(xComment))

    End Function

    Private Shared Function GetAvatar(ByVal sAvatar As String, ByVal sModulus As String) As String

        If sAvatar Is Nothing Then Return DefaultAvatar(sModulus)
        If Len(sAvatar) = 0 Then Return DefaultAvatar(sModulus)

        Try

            Dim sExtension As String = ""

            Dim Bm As New Bitmap(New MemoryStream(Convert.FromBase64String(sAvatar)))
            Dim Bf As ImageFormat = Bm.RawFormat

            If (Bf.Equals(System.Drawing.Imaging.ImageFormat.Jpeg)) Then
                sExtension = "jpeg"
            End If

            If (Bf.Equals(System.Drawing.Imaging.ImageFormat.Bmp)) Then
                sExtension = "bmp"
            End If

            If (Bf.Equals(System.Drawing.Imaging.ImageFormat.Gif)) Then
                sExtension = "gif"
            End If

            If (Bf.Equals(System.Drawing.Imaging.ImageFormat.Png)) Then
                sExtension = "png"
            End If

            If Len(sExtension) = 0 Then Return DefaultAvatar(sModulus)

            Return "data:image/" & sExtension & ";base64," & sAvatar

        Catch
        End Try

        Return DefaultAvatar(sModulus)

    End Function

    Private Shared Function DefaultAvatar(ByVal sModulus As String) As String

        Return "http://www.gravatar.com/avatar/" & MakeMD5(sModulus) & "?s=32&d=identicon"

    End Function

    Private Shared Function MakeDesc(ByVal xSpot As SpotEx) As String

        Dim sDesc As String = Replace(Trim(xSpot.Body), "[br]", vbCrLf, , , CompareMethod.Text)

        Do While sDesc.StartsWith(vbCrLf)
            sDesc = sDesc.Substring(2)
        Loop

        Do While sDesc.EndsWith(vbCrLf)
            sDesc = sDesc.Substring(0, sDesc.Length - 2)
        Loop

        sDesc = Replace$(sDesc.Replace(vbCrLf, "[br]"), "&#", "[aa]") ' allow special chars

        If Not xSpot.OldInfo Is Nothing Then

            sDesc = Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(sDesc, "&amp;lt;br />", "[br]", , , CompareMethod.Text), "&lt;br /&gt;", "[br]", , , CompareMethod.Text), "&quot;", Chr(34), , , CompareMethod.Text), "&amp;", "&", , , CompareMethod.Text), "<br>", "[br]", , , CompareMethod.Text), "<br/>", "[br]", , , CompareMethod.Text), "<br />", "[br]", , , CompareMethod.Text), "</br>", "", , , CompareMethod.Text)

        End If

        sDesc = HtmlEncode(sDesc)

        sDesc = Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(sDesc, "[b]", "<b>", , , CompareMethod.Text), "[/b]", "</b>", , , CompareMethod.Text), "[u]", "<u>", , , CompareMethod.Text), "[/u]", "</u>", , , CompareMethod.Text), "[i]", "<i>", , , CompareMethod.Text), "[/i]", "</i>", , , CompareMethod.Text), "[br]", "<br>", , , CompareMethod.Text), "[aa]", "&#")

        If (sDesc.IndexOf("<br><br><br>")) > 250 Then
            sDesc = Replace$(sDesc, "<br><br><br>", "<br clear='all'>&nbsp;<br>")
        End If

        sDesc = InsertSmileys(Linkify(sDesc))

        Return sDesc & "</b></b></b></i></i></u></u>"

    End Function

    Private Shared Function MakeBody(ByVal xComment As Comment) As String

        Dim sDesc As String = Replace(Trim(xComment.Body), "[br]", vbCrLf, , , CompareMethod.Text)

        Do While sDesc.StartsWith(vbCrLf)
            sDesc = sDesc.Substring(2)
        Loop

        Do While sDesc.EndsWith(vbCrLf)
            sDesc = sDesc.Substring(0, sDesc.Length - 2)
        Loop

        Dim zBod As String = sDesc.Replace(vbTab, " ").Replace(vbCrLf, vbTab)
        Dim zSplit() As String = zBod.Split(vbTab.ToCharArray())
        If UBound(zSplit) > 30 Then sDesc = zBod.Replace(vbTab, " ")

        sDesc = Replace$(sDesc.Replace(vbCrLf, "[br]"), "&#", "[aa]", , , CompareMethod.Text) ' allow special chars

        sDesc = HtmlEncode(sDesc)

        sDesc = Replace$(Replace$(sDesc, "[br]", "<br>", , , CompareMethod.Text), "[aa]", "&#")

        sDesc = InsertSmileys(Linkify(sDesc))

        Return sDesc

    End Function

    Friend Shared Function MakeTable(ByVal xSpot As SpotEx) As String

        Dim xTable As New StringBuilder

        xTable.Append("<TABLE BORDER=0>")

        Dim sGroups As String = ""
        Dim TransCat As String = ""

        If Not xSpot.OldInfo Is Nothing Then sGroups = xSpot.OldInfo.Groups

        xTable.Append(GetCatInfo(xSpot.Category, xSpot.SubCats, xSpot.Tag, sGroups, xSpot.Poster, xSpot.User))

        If xSpot.OldInfo Is Nothing Then
            xTable.Append("<TR><TD><b>Omvang</b></TD><TD>" & ConvertSize(xSpot.Filesize) & "</TD></TR>")
            xTable.Append("<TR><TD>&nbsp;</TD><TD>&nbsp;</TD></TR>")
        End If

        If Len(xSpot.Web) > 0 Then
            xTable.Append("<TR><TD><b>Website</b></TD><TD><A onfocus='this.blur()' TITLE='Link openen' HREF=" & Chr(34) & "link:" & SafeHref(xSpot.Web) & Chr(34) & ">" & HtmlEncode(URLDecode(xSpot.Web)) & "</A></TD></TR>")
        Else
            xTable.Append("<TR><TD><b>Website</b></TD><TD><A onfocus='this.blur()' TITLE='Link openen' HREF=" & Chr(34) & "link:" & SafeHref(MakeGoogleSearch(xSpot.Title)) & Chr(34) & ">" & HtmlEncode(URLDecode(MakeGoogleSearch(xSpot.Title))) & "</A></TD></TR>")
        End If

        If Not xSpot.OldInfo Is Nothing Then
            xTable.Append("<TR><TD>&nbsp;</TD><TD>&nbsp;</TD></TR>")
            xTable.Append("<TR><TD><b>Bestandsnaam&nbsp;&nbsp;&nbsp;</b></TD><TD><B><A onfocus='this.blur()' TITLE='Zoeken naar NZB' HREF='link:" & MakeNZBSearch(HtmlDecode(xSpot.OldInfo.FileName)) & "'>" & HtmlEncode(xSpot.OldInfo.FileName) & "</A></B></TD></TR>")
        End If

        xTable.Append(sGroups)
        xTable.Append("</TABLE>")

        Return xTable.ToString

    End Function

    Private Shared Function MakeCats(ByVal hCat As Byte, ByVal sCat As Char, ByVal TheCats As Dictionary(Of String, String), ByVal sColor As String, ByVal sColor2 As String) As String

        Dim sApp As String
        Dim xRet As New StringBuilder

        If TheCats.Count > 0 Then
            sApp = vbNullString
            xRet.Append("<TR><TD><b>" & Utils.TranslateCatDesc(hCat, sCat) & "</b>&nbsp;&nbsp;&nbsp;</TD><TD>")
            For Each sG As String In TheCats.Keys
                If hCat = 1 Or hCat = 6 Then
                    xRet.Append(sApp & CreateClassLink("query:" & URLEncode(TheCats.Item(sG) & "_" & "cats MATCH '" & StripNonAlphaNumericCharacters("1" & sG) & " OR " & StripNonAlphaNumericCharacters("6" & sG) & "'"), TheCats.Item(sG), "category", "Zoeken") & "</TD></TR>")
                Else
                    xRet.Append(sApp & CreateClassLink("query:" & URLEncode(TheCats.Item(sG) & "_" & "cats MATCH '" & StripNonAlphaNumericCharacters(CStr(hCat) & sG) & "'"), TheCats.Item(sG), "category", "Zoeken") & "</TD></TR>")
                End If
                sApp = "<TR>&nbsp;<TD></TD><TD>"
            Next
        End If

        Return xRet.ToString

    End Function

    Private Shared Function MakeNZBSearch(ByVal sTitle As String) As String

        Return "http://binsearch.info/index.php?q=" & URLEncode(sTitle) & HtmlEncode("&max=250&adv_age=&server=")

    End Function

    Private Shared Function CreateClassLink(ByVal sTarget As String, ByVal sText As String, ByVal sClass As String, Optional ByVal sTooltip As String = "") As String

        Return "<A CLASS='" & sClass & "' " & sIIF(Len(sTooltip) > 0, " TITLE='" & sTooltip.Replace("'", "''") & "'", " ") & " onfocus='this.blur()' HREF='" & sTarget.Replace("'", "''") & "'>" & HtmlEncode(sText) & "</A>"

    End Function

    Private Shared Function GetCatInfo(ByVal hCat As Byte, ByVal Cats As String, ByVal zTags As String, ByRef zGroups As String, ByVal sPoster As String, ByVal zUser As UserInfo) As String

        Dim zAdd As String
        Dim zCat() As String
        Dim TypeCat As Byte = 0
        Dim sColor As String = "black"
        Dim sColor2 As String = "blue"

        Dim sOutTags As String = vbNullString
        Dim OutPoster As String = vbNullString

        Dim iCatA As New Dictionary(Of String, String)
        Dim iCatB As New Dictionary(Of String, String)
        Dim iCatC As New Dictionary(Of String, String)
        Dim iCatD As New Dictionary(Of String, String)
        Dim iCatZ As New Dictionary(Of String, String)

        Dim CatTags As New Collection
        Dim CatGroups As New Collection

        zCat = Split(Cats, "|")

        For i = 0 To UBound(zCat)
            If Len(zCat(i)) > 0 Then

                zAdd = TranslateCat(hCat, zCat(i))

                If Len(zAdd) = 0 Then Continue For

                Select Case Microsoft.VisualBasic.Left(zCat(i), 1).ToLower

                    Case "a"

                        If Not iCatA.ContainsKey(zCat(i)) Then
                            iCatA.Add(zCat(i), zAdd)
                        End If

                    Case "b"

                        If Not iCatB.ContainsKey(zCat(i)) Then
                            iCatB.Add(zCat(i), zAdd)
                        End If

                    Case "c"

                        If Not iCatC.ContainsKey(zCat(i)) Then
                            iCatC.Add(zCat(i), zAdd)
                        End If

                    Case "d"

                        If hCat = 9 Then

                            Select Case zCat(i).ToLower
                                Case "d75", "d74", "d73", "d72"
                                    If Cats.Contains("d2") Then
                                        Continue For ' Double Cat fix
                                    End If
                            End Select

                        End If

                        If Not iCatD.ContainsKey(zCat(i)) Then
                            iCatD.Add(zCat(i), zAdd)
                        End If

                    Case "z"

                        If Not iCatZ.ContainsKey(zCat(i)) Then

                            TypeCat = CByte(Val(zCat(i).ToLower.Replace("z", "")) + 1)
                            iCatZ.Add(zCat(i), zAdd)

                        End If

                End Select
            End If
        Next

        zCat = Split(zTags, " ")

        For i = 0 To UBound(zCat)
            If Len(Trim$(zCat(i))) > 0 Then
                CatTags.Add(zCat(i))
            End If
        Next

        zCat = Split(zGroups, "|")

        For i = 0 To UBound(zCat)
            If Len(Trim$(zCat(i))) > 0 Then
                CatGroups.Add(zCat(i))
            End If
        Next

        Dim sApp As String

        If CatTags.Count > 0 Then
            sApp = vbNullString
            sOutTags = "<TR><TD><b>Tag" & sIIF(CatTags.Count > 1, "s", "") & "</b></TD><TD>"
            For Each sG As String In CatTags
                If sG <> sPoster Then
                    sOutTags += sApp & CreateClassLink("query:" & URLEncode(StripNonAlphaNumericCharacters(sG) & "_" & "tag MATCH '" & StripNonAlphaNumericCharacters(sG).ToLower & "'"), StripNonAlphaNumericCharacters(sG), "category", "Zoeken") & "</TD></TR>"
                    sApp = "<TR>&nbsp;<TD></TD><TD>"
                End If
            Next
            If Len(sApp) = 0 Then sOutTags = vbNullString
        End If

        If CatGroups.Count > 0 Then
            sApp = vbNullString
            zGroups = "<TR><TD><b>Nieuwsgroep" & sIIF(CatGroups.Count > 1, "en", "") & "</b></TD><TD>"
            For Each sG As String In CatGroups
                If sG <> "Other" Then
                    zGroups += sApp & HtmlEncode(sG) & "</TD></TR>"
                    sApp = "<TR><TD>&nbsp;</TD><TD>"
                End If
            Next
            If Len(sApp) = 0 Then zGroups = vbNullString
        Else
            zGroups = vbNullString
        End If

        Dim sTooltip As String = "Onbekend"

        If zUser.ValidSignature Then
            sTooltip = HtmlEncode(MakeUnique(zUser.Modulus))
        End If

        If Len(zUser.Organisation) > 0 Then
            sTooltip += vbCrLf & HtmlEncode(zUser.Organisation)
        End If

        Dim SuperSpot As Boolean = zUser.ValidSignature And WhiteList.Contains(zUser.Modulus)

        OutPoster = "<TR><TD><b>Afzender</b></TD><TD>" & CreateClassLink("menu:" & sIIF(zUser.ValidSignature, zUser.Modulus, "") & "_" & StripNonAlphaNumericCharacters(sPoster), StripNonAlphaNumericCharacters(sPoster), sIIF(SuperSpot, "trusted", "from"), sTooltip) & "</TD></TR>"

        Dim sBr As String = "<TR><TD>&nbsp;</TD><TD>&nbsp;</TD></TR>"

        zGroups += sBr & OutPoster & sOutTags

        Dim TransCat As String = CatDesc(hCat, TypeCat)
        Dim sSort As String = "<TR><TD><b>Categorie&nbsp;&nbsp;&nbsp;</b></TD><TD>" & CreateClassLink("query:" & URLEncode(TransCat & "_" & "cat = " & CStr(hCat)), TransCat, "category", "Zoeken") & "</TD></TR>"

        Dim sOutZ As String = MakeCats(hCat, "z"c, iCatZ, sColor, sColor2)
        Dim sOutA As String = MakeCats(hCat, "a"c, iCatA, sColor, sColor2)
        Dim sOutB As String = MakeCats(hCat, "b"c, iCatB, sColor, sColor2)
        Dim sOutC As String = MakeCats(hCat, "c"c, iCatC, sColor, sColor2)
        Dim sOutD As String = MakeCats(hCat, "d"c, iCatD, sColor, sColor2)

        Select Case hCat

            Case 2

                If TypeCat > 1 Then sSort = sOutZ
                sSort += sOutA & sOutB & sOutC & sOutD

            Case 3
                sSort += sOutB & sOutA & sOutC

            Case 4
                sSort += sOutA & sOutB

            Case Else
                sSort += sOutA & sOutB & sOutC & sOutD

        End Select

        Return sSort

    End Function

    Friend Shared Function InsertSmileys(ByVal sHTML As String) As String

        Do While sHTML.Contains("[img=")

            Dim sp As Integer = sHTML.IndexOf("[img=")
            Dim sp2 As Integer = sHTML.IndexOf("]", sp)

            If (sp >= 0 And sp2 > 0) And ((sp2 - sp) > 5) Then

                Dim sSmile As String = StripNonAlphaNumericCharacters(sHTML.Substring(sp + 5, sp2 - (sp + 5)))

                sHTML = sHTML.Remove(sp, sp2 - (sp - 1))

                Dim sFile As String = SettingsFolder() & "\Images\Smileys\" & sSmile & ".gif"
                If FileExists(sFile) Then sHTML = sHTML.Insert(sp, "<IMG onfocus='this.blur()' title=" & Chr(34) & sSmile & Chr(34) & " SRC=" & Chr(34) & sFile & Chr(34) & ">")

            Else
                Exit Do
            End If

        Loop

        Return sHTML

    End Function

    Friend Shared Function Linkify(ByVal sIn As String) As String

        Static _Linkify As New Regex("\(?\b(https|http)://[-A-Za-z0-9+&@#/%?=~_()|!:,.;]*[-A-Za-z0-9+&@#/%=~_()|]", RegexOptions.IgnoreCase)

        Dim Matches As MatchCollection = _Linkify.Matches(sIn)

        For Each xMatch As Match In Matches

            sIn = sIn.Replace(xMatch.Value, "<A onfocus='this.blur()' TITLE='Link openen' HREF=" & Chr(34) & "link:" & SafeHref(HtmlDecode(xMatch.Value)) & Chr(34) & ">" & xMatch.Value & "</A>")

        Next

        Return sIn

    End Function

    Private Shared Function MakeGoogleSearch(ByVal sTitle As String) As String

        Return "http://www.google.nl/search?q=" & URLEncode(sTitle)

    End Function

    Private Shared Function SHA1(ByVal sVal As String) As String

        Dim Kl As New SHA1Managed
        Return StripNonAlphaNumericCharacters(Convert.ToBase64String(Kl.ComputeHash(MakeLatin(sVal))))

    End Function

End Class
