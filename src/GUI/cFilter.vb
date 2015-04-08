Friend Class cFilter

    Private Const sfName As String = "filters.xml"
    Friend FilterOverview As New List(Of Spotnet.FilterCat)

    Public Function DefaultFilters(ByRef zErr As String) As Boolean

        Dim LF As New List(Of Spotnet.FilterCat)
        Dim lMarg As Long = 50

        LF.Add(CreateFilter(0, CatDesc(1), "cat = 1", "\Images\video2.ico"))
        LF.Add(CreateFilter(0, CatDesc(6), "cat = 6", "\Images\series2.ico"))
        LF.Add(CreateFilter(0, CatDesc(5), "cat = 5", "\Images\books2.ico"))
        LF.Add(CreateFilter(0, CatDesc(2), "cat = 2", "\Images\audio2.ico"))
        LF.Add(CreateFilter(0, CatDesc(3), "cat = 3", "\Images\games2.ico"))
        LF.Add(CreateFilter(0, CatDesc(4), "cat = 4", "\Images\applications2.ico"))
        LF.Add(CreateFilter(0, CatDesc(9), "cat = 9", "\Images\x2.ico"))

        LF.Add(CreateFilter(22, "Nieuw", "rowid > [SN:NEW]", "1.ico", "2.ico", False))
        LF.Add(CreateFilter(22, "Vandaag", "date > ( [SN:DATE] - 86400 )", "1.ico", "2.ico", False))

        FilterOverview.Clear()
        Return AddFilters(LF, True, zErr)

    End Function

    Public Function LoadFilters() As Boolean

        Dim root As Xml.XmlElement
        Dim doc As New Xml.XmlDocument
        Dim child2 As Xml.XmlElement
        Dim FilterX As New HashSet(Of String)

        Try

            If Not System.IO.File.Exists(SettingsFolder() & "\" & sfName) Then
                DefaultFilters("")
            End If

            doc.XmlResolver = Nothing
            doc.Load(SettingsFolder() & "\" & sfName)
            root = doc.DocumentElement

            If (root.ChildNodes.Count) < 1 Then
                DefaultFilters("")
                doc.Load(SettingsFolder() & "\" & sfName)
                root = doc.DocumentElement
            End If

            FilterOverview = New List(Of Spotnet.FilterCat)

            For Each child2 In root

                Dim TheFil As New Spotnet.FilterCat

                With TheFil

                    .sName = child2.GetAttribute("Name")

                    If Not child2.GetAttribute("Visible") Is Nothing Then
                        If LCase(child2.GetAttribute("Visible")) = "false" Then Continue For
                    End If

                    If Not child2.GetAttribute("Image") Is Nothing Then .sImage = child2.GetAttribute("Image")
                    If Not child2.GetAttribute("SelectedImage") Is Nothing Then .sSelected = child2.GetAttribute("SelectedImage")
                    If Not child2.GetAttribute("Margin") Is Nothing Then .sMargin = CStr(Val(child2.GetAttribute("Margin")))

                    If Len(.sImage) > 0 Then
                        If Not .sImage.Contains(":") Then
                            .sImage = SettingsFolder() & .sImage
                        End If
                    End If

                    If Len(.sSelected) > 0 Then
                        If Not .sSelected.Contains(":") Then
                            .sSelected = SettingsFolder() & .sSelected
                        End If
                    End If

                    .sQuery = child2.InnerText.Trim
                    .sQuery = Strings.Replace(.sQuery, "cat = 1 AND cats MATCH '1b4 OR 1d11'", "cat = 6", 1, , CompareMethod.Text)

                    If .sQuery.Contains("SCat =") Then Continue For
                    If .sQuery.Contains("topcat =") Then Continue For
                    If .sQuery.Contains("subcat IN") Then Continue For
                    If .sQuery.Contains("subcat =") Then Continue For
                    If .sQuery.Contains("subcats LIKE") Then Continue For
                    If .sQuery.Contains("subcats like") Then Continue For

                    If .sQuery.Contains("tag = '") Then
                        .sQuery = .sQuery.ToLower.Replace("tag = '", "tag MATCH '")
                    End If

                    If .sQuery.Contains("sender = '") Then
                        .sQuery = .sQuery.ToLower.Replace("sender = '", "sender MATCH '")
                    End If

                    If Len(.sImage) = 0 And Len(.sName) > 0 Then
                        If .sQuery.ToLower.Contains("tag match '") Or .sQuery.ToLower.Contains("sender match '") Then
                            .sImage = SettingsFolder() & "\Images\people2.ico"
                        Else
                            If .sQuery.ToLower.Contains("tag like") Or .sQuery.ToLower.Contains("tag = '") Then
                                .sImage = SettingsFolder() & "\Images\tag2.ico"
                            Else
                                .sImage = SettingsFolder() & "\Images\custom2.ico"
                            End If
                        End If
                    End If

                    If FilterX.Add(.sName) Then FilterOverview.Add(TheFil)

                End With

            Next

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Public Sub RemoveFilter(ByVal fName As String)

        Dim lp As Spotnet.FilterCat
        Dim doc As New Xml.XmlDocument

        doc.XmlResolver = Nothing

        If Not FilterExist(fName) Then Exit Sub

        If System.IO.File.Exists(SettingsFolder() & "\" & sfName) Then
            doc.Load(SettingsFolder() & "\" & sfName)
        Else
            Exit Sub
        End If

        Dim child2 As Xml.XmlElement

        For Each child2 In doc.DocumentElement
            If UCase$(child2.GetAttribute("Name")) = UCase$(fName) Then
                doc.DocumentElement.RemoveChild(child2)
            End If
        Next

        For Each lp In FilterOverview
            If Trim$(UCase$(lp.Name)) = Trim$(UCase$(fName)) Then
                FilterOverview.Remove(lp)
                Exit For
            End If
        Next

        If FileExists(SettingsFolder() & "\" & sfName) Then
            System.IO.File.SetAttributes(SettingsFolder() & "\" & sfName, IO.FileAttributes.Normal)
        End If

        doc.Save(SettingsFolder() & "\" & sfName)

    End Sub

    Public Sub SwapFilter(ByVal fName As String, ByVal bUp As Boolean)

        Dim doc As New Xml.XmlDocument

        doc.XmlResolver = Nothing

        If Not FilterExist(fName) Then Exit Sub

        If System.IO.File.Exists(SettingsFolder() & "\" & sfName) Then
            doc.Load(SettingsFolder() & "\" & sfName)
        Else
            Exit Sub
        End If

        Dim child2 As Xml.XmlNode
        Dim xchild As Xml.XmlNode

        For Each child2 In doc.DocumentElement
            If UCase$(CType(child2, Xml.XmlElement).GetAttribute("Name")) = UCase$(fName) Then
                If bUp Then
                    xchild = child2.PreviousSibling
                    doc.DocumentElement.RemoveChild(xchild)
                    doc.DocumentElement.InsertAfter(xchild, child2)
                Else
                    xchild = child2.NextSibling
                    doc.DocumentElement.RemoveChild(child2)
                    doc.DocumentElement.InsertAfter(child2, xchild)
                End If
            End If
        Next

        If FileExists(SettingsFolder() & "\" & sfName) Then
            System.IO.File.SetAttributes(SettingsFolder() & "\" & sfName, IO.FileAttributes.Normal)
        End If

        doc.Save(SettingsFolder() & "\" & sfName)

    End Sub

    Private Function CreateFilter(ByVal Margin As Long, ByVal sName As String, ByVal sQuery As String, Optional ByVal sImage As String = "", Optional ByVal Selected As String = "", Optional ByVal IsVisible As Boolean = True) As Spotnet.FilterCat

        Dim ZL As New Spotnet.FilterCat

        ZL.sName = sName
        ZL.sQuery = sQuery
        ZL.sImage = sImage
        ZL.Visible = IsVisible
        ZL.sSelected = Selected
        ZL.sMargin = CStr(Margin) & ",0,0,0"

        Return ZL

    End Function

    Public Function AddFilter(ByVal sName As String, ByVal sQuery As String, Optional ByVal sImage As String = "", Optional ByRef sError As String = "") As Boolean

        Dim z2 As New List(Of Spotnet.FilterCat)
        z2.Add(CreateFilter(0, sName, sQuery, sImage))
        Return AddFilters(z2, False, sError)

    End Function

    Friend Function FilterExist(ByVal sName As String) As Boolean

        Dim lp As Spotnet.FilterCat

        If Len(Trim$(sName)) > 0 Then
            For Each lp In FilterOverview
                If Trim$(UCase$(lp.Name)) = Trim$(UCase$(sName)) Then
                    Return True
                End If
            Next
            Return False
        Else
            Return True
        End If

    End Function

    Public Function AddFilters(ByVal zFilter As List(Of Spotnet.FilterCat), ByVal bClean As Boolean, Optional ByRef sError As String = "") As Boolean

        Dim root As Xml.XmlElement
        Dim doc As New Xml.XmlDocument

        doc.XmlResolver = Nothing

        For Each ZZ As Spotnet.FilterCat In zFilter
            If FilterExist(ZZ.Name) Then
                sError = "Filternaam bestaat al!"
                Return False
            End If
        Next

        If System.IO.File.Exists(SettingsFolder() & "\" & sfName) And (Not bClean) Then
            doc.Load(SettingsFolder() & "\" & sfName)
            root = doc.DocumentElement
        Else
            root = doc.CreateElement(Spotname)
        End If

        For Each ZZ As Spotnet.FilterCat In zFilter
            Dim child As Xml.XmlElement = doc.CreateElement("Filter")
            child.SetAttribute("Name", ZZ.Name)
            If Not ZZ.Visible Then child.SetAttribute("Visible", "false")
            If Len(ZZ.NormalImage) > 0 Then child.SetAttribute("Image", ZZ.NormalImage.Replace(SettingsFolder, ""))
            If Len(ZZ.SelectedImage) > 0 Then child.SetAttribute("SelectedImage", ZZ.SelectedImage.Replace(SettingsFolder, ""))
            If Len(ZZ.Margin) > 0 And ZZ.Margin <> "0,0,0,0" And ZZ.Margin <> "0" Then child.SetAttribute("Margin", ZZ.Margin)
            child.InnerXml = "<![CDATA[" & sIIF(Len(Trim$(ZZ.Query)) = 0, " ", ZZ.Query) & "]]>"
            root.AppendChild(child)
        Next

        doc.AppendChild(root)

        If FileExists(SettingsFolder() & "\" & sfName) Then
            System.IO.File.SetAttributes(SettingsFolder() & "\" & sfName, IO.FileAttributes.Normal)
        End If

        doc.Save(SettingsFolder() & "\" & sfName)

        Return True

    End Function

End Class
