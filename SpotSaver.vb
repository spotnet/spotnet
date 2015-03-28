Imports System.IO
Imports System.Text
Imports System.Data.Common

Friend Class SpotSaver

    Private Db As SqlDB

    Public Event ProgressChanged(ByVal sMsg As String, ByVal lVal As Integer)

    Public Function Connect(ByVal dbFile As String, ByRef sError As String) As Boolean

        Try

            Close()
            Db = New SqlDB

            If Not Db.Connect(dbFile, False) Then Throw New Exception("iConnect")

            Return True

        Catch ex As AccessViolationException

            sError = ex.Message
            Return False

        Catch ex As Exception

            sError = "iConnect: " & ex.Message
            Return False

        End Try

    End Function

    Function InsertSpots(ByVal zList As List(Of Spot), ByVal DeleteList As HashSet(Of String), ByRef xSave As rSaveSpots, ByRef sErr As String) As Boolean

        Dim ReturnParam As New rSaveSpots

        Try

            xSave = Nothing

            If (zList.Count = 0) And (DeleteList.Count = 0) Then
                Close()
                Return True
            End If

            If Not Db.ExecuteNonQuery("PRAGMA synchronous = OFF", "") = 0 Then Throw New Exception("PRAGMA synchronous")
            If Not Db.ExecuteNonQuery("PRAGMA journal_mode = MEMORY", "") = 0 Then Throw New Exception("PRAGMA journal_mode")
            If Not Db.ExecuteNonQuery("PRAGMA temp_store = MEMORY;", "") = 0 Then Throw New Exception("PRAGMA temp_store")
            If Not Db.ExecuteNonQuery("PRAGMA locking_mode = EXCLUSIVE", "") = 0 Then Throw New Exception("PRAGMA locking_mode")
            If Not Db.ExecuteNonQuery("PRAGMA cache_size = " & CStr(My.Settings.DatabaseCache), "") = 0 Then Throw New Exception("PRAGMA cache_size")

            If zList.Count > 0 Then

                AddSpots(zList, DeleteList, ReturnParam)

            End If

            If DeleteList.Count > 0 And My.Settings.DatabaseCount > 0 Then DeleteSpots(DeleteList, ReturnParam)

            RaiseEvent ProgressChanged("Indexen bijwerken...", 0)

            My.Settings.DatabaseMax = Db.ExecuteScalar("SELECT MAX(rowid) FROM spots", "")
            My.Settings.DatabaseCount = Db.ExecuteScalar("SELECT COUNT(*) FROM spots", "")
            My.Settings.DatabaseFilter = My.Settings.DatabaseCount - Math.Abs(Db.ExecuteScalar("SELECT COUNT(*) FROM spots WHERE cat=9", ""))

            My.Settings.Save()

            RaiseEvent ProgressChanged("Database sluiten...", 0)

            Close()

            xSave = ReturnParam
            Return True

        Catch ex As Exception

            Close()
            sErr = "InsertSpots: " & ex.Message
            Return False

        End Try

    End Function

    Private Function DeleteSpots(ByVal DeleteList As HashSet(Of String), ByVal xSave As rSaveSpots) As Boolean

        Dim zPos As Integer = 0
        Dim LastP As Integer = -1
        Dim ProgressValue As Integer = 0

        If (DeleteList.Count = 0) Then Return True

        RaiseEvent ProgressChanged("Spots verwijderen...", 0)

        Dim IDList As List(Of Long) = MessageToRow(DeleteList)

        If (IDList.Count = 0) Then Return True

        Dim dbTrans As DbTransaction = Db.BeginTransaction
        Dim dbCmd As DbCommand = Db.CreateCommand

        dbCmd.CommandText = "DELETE FROM search WHERE docid=?; DELETE FROM spots WHERE rowid=?;"

        Dim Param1 As DbParameter = dbCmd.CreateParameter()
        Dim Param2 As DbParameter = dbCmd.CreateParameter()

        dbCmd.Parameters.Add(Param1)
        dbCmd.Parameters.Add(Param2)

        For Each xDelID As Long In IDList

            zPos += 1
            ProgressValue = CInt((100 / IDList.Count) * zPos)

            If ProgressValue <> LastP Then
                RaiseEvent ProgressChanged("", ProgressValue)
                LastP = ProgressValue
            End If

            Param1.Value = xDelID
            Param2.Value = xDelID

            If (dbCmd.ExecuteNonQuery() > 0) Then
                xSave.SpotsDeleted += 1
            End If

        Next

        RaiseEvent ProgressChanged("", 0)

        dbTrans.Commit()

        dbTrans.Dispose()
        dbCmd.Dispose()

        Return True

    End Function

    Friend Function AddComments(ByRef zList As List(Of Comment), ByRef xSave As rSaveComments, ByRef sErr As String) As Boolean

        Dim zPos As Integer = 0
        Dim LastP As Integer = -1
        Dim lRet As Integer = -1
        Dim ProgressValue As Integer = 0
        Dim ReturnParam As New rSaveComments
        Dim AtPos As Integer = -1

        Try

            xSave = Nothing

            If zList.Count = 0 Then
                Close()
                Return False
            End If

            RaiseEvent ProgressChanged("Reacties opslaan...", 0)

            If Not Db.ExecuteNonQuery("PRAGMA page_size = 4096;", "") = 0 Then Throw New Exception("PRAGMA page_size")
            If Not Db.ExecuteNonQuery("PRAGMA synchronous = OFF", "") = 0 Then Throw New Exception("PRAGMA synchronous")
            If Not Db.ExecuteNonQuery("PRAGMA journal_mode = MEMORY", "") = 0 Then Throw New Exception("PRAGMA journal_mode")
            If Not Db.ExecuteNonQuery("PRAGMA temp_store = MEMORY;", "") = 0 Then Throw New Exception("PRAGMA temp_store")
            If Not Db.ExecuteNonQuery("PRAGMA locking_mode = EXCLUSIVE", "") = 0 Then Throw New Exception("PRAGMA locking_mode")
            If Not Db.ExecuteNonQuery("PRAGMA cache_size = " & CStr(My.Settings.DatabaseCache), "") = 0 Then Throw New Exception("PRAGMA cache_size")

            Dim dbTrans As DbTransaction = Db.BeginTransaction

            Db.ExecuteNonQuery("CREATE VIRTUAL TABLE comments USING fts4(spot TEXT,matchinfo=fts3)", "")

            Dim dbCmd As DbCommand = Db.CreateCommand
            dbCmd.CommandText = "INSERT OR IGNORE INTO comments(docid, spot) VALUES (?,?);"

            Dim Param1 As DbParameter = dbCmd.CreateParameter()
            dbCmd.Parameters.Add(Param1)

            Dim Param2 As DbParameter = dbCmd.CreateParameter()
            dbCmd.Parameters.Add(Param2)

            For Each CurCom As Comment In zList

                zPos += 1
                ProgressValue = CInt((100 / zList.Count) * zPos)

                If ProgressValue <> LastP Then
                    RaiseEvent ProgressChanged("", ProgressValue)
                    LastP = ProgressValue
                End If

                AtPos = CurCom.MessageID.IndexOf("@")
                If AtPos < 2 Then Continue For

                Param1.Value = CurCom.Article
                Param2.Value = CurCom.MessageID.Substring(0, AtPos)

                lRet = dbCmd.ExecuteNonQuery()

                If lRet < 0 Then
                    Throw New Exception("Comments.ExecuteNonQuery")
                End If

                If lRet > 0 Then
                    ReturnParam.CommentsAdded += 1
                End If

            Next

            RaiseEvent ProgressChanged("", 0)

            dbTrans.Commit()

            dbTrans.Dispose()
            dbCmd.Dispose()

            RaiseEvent ProgressChanged("Database sluiten...", 0)

            Close()

            xSave = ReturnParam
            Return True

        Catch ex As Exception

            Close()

            sErr = "InsertComments: " & ex.Message
            Return False

        End Try

    End Function

    Private Function AddSpots(ByVal SpotList As List(Of Spot), ByVal DeleteList As HashSet(Of String), ByVal xSave As rSaveSpots) As Boolean

        Dim sCats As String
        Dim zPos As Integer = 0
        Dim lRet As Integer = -1
        Dim LastP As Integer = -1
        Dim ProgressValue As Integer = 0
        Dim xNull As DBNull = DBNull.Value
        Dim bCountCats As Boolean = SpotList.Count < 10000

        If (SpotList.Count = 0) Then Return True

        RaiseEvent ProgressChanged("Spots opslaan...", 0)

        Dim dbTrans As DbTransaction = Db.BeginTransaction
        Dim dbCmd As DbCommand = Db.CreateCommand

        dbCmd.CommandText = "INSERT OR IGNORE INTO spots(rowid,key,cat,subcat,extcat,date,filesize,cats,sender,tag,subject,msgid,modulus) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?);"

        Dim Param(12) As DbParameter

        For xl As Integer = 0 To UBound(Param)
            Param(xl) = dbCmd.CreateParameter()
            dbCmd.Parameters.Add(Param(xl))
        Next

        Dim AddList As New List(Of Spot)

        For Each xSpot As Spot In SpotList

            zPos += 1
            ProgressValue = CInt((50 / SpotList.Count) * zPos)

            If ProgressValue <> LastP Then
                RaiseEvent ProgressChanged("", ProgressValue)
                LastP = ProgressValue
            End If

            With xSpot

                If DeleteList.Contains(.MessageID) Then
                    DeleteList.Remove(.MessageID)
                    Continue For
                End If

                If bCountCats Then
                    xSave.NewCats(.Category) += 1
                End If

                Param(0).Value = .Article
                Param(1).Value = .KeyID
                Param(2).Value = .Category
                Param(3).Value = CInt((.Category * 100) + .SubCat)
                Param(4).Value = CInt((.Category * 100) + TranslateInfo(.Category, .SubCats))
                Param(5).Value = .Stamp
                Param(6).Value = .Filesize

                sCats = CStr(.Category) & .SubCats.Replace("|", " " & CStr(.Category))
                .SubCats = CStr(.Category) & " " & sCats.Substring(0, sCats.Length - 2)

                Param(7).Value = .SubCats
                Param(8).Value = .Poster

                If Not .Tag Is Nothing Then
                    Param(9).Value = .Tag
                Else
                    Param(9).Value = xNull
                End If

                Param(10).Value = .Title
                Param(11).Value = .MessageID

                If Not .Modulus Is Nothing Then
                    Param(12).Value = .Modulus
                Else
                    Param(12).Value = xNull
                End If

                lRet = dbCmd.ExecuteNonQuery()

                If lRet < 0 Then
                    Throw New Exception("Spots.ExecuteNonQuery.1")
                End If

                If lRet > 0 Then AddList.Add(xSpot)

            End With
        Next

        Db.ExecuteNonQuery("CREATE INDEX IF NOT EXISTS catidx ON spots(cat)", "")

        zPos = 0

        dbCmd = Db.CreateCommand
        dbCmd.CommandText = "INSERT OR IGNORE INTO search(docid,cats,sender,tag,subject) VALUES (?,?,?,?,?);"

        ReDim Param(4)

        For xl As Integer = 0 To UBound(Param)
            Param(xl) = dbCmd.CreateParameter()
            dbCmd.Parameters.Add(Param(xl))
        Next

        For Each xSpot As Spot In AddList

            zPos += 1
            ProgressValue = CInt((50 / AddList.Count) * zPos) + 50

            If ProgressValue <> LastP Then
                RaiseEvent ProgressChanged("", ProgressValue)
                LastP = ProgressValue
            End If

            With xSpot

                Param(0).Value = .Article
                Param(1).Value = .SubCats
                Param(2).Value = .Poster

                If Not .Tag Is Nothing Then
                    Param(3).Value = .Tag
                Else
                    Param(3).Value = xNull
                End If

                Param(4).Value = .Title

                lRet = dbCmd.ExecuteNonQuery()

                If lRet < 0 Then
                    Throw New Exception("Spots.ExecuteNonQuery.2")
                End If

                If lRet > 0 Then xSave.SpotsAdded += 1

            End With
        Next

        RaiseEvent ProgressChanged("", 0)

        dbTrans.Commit()

        dbTrans.Dispose()
        dbCmd.Dispose()

        Return True

    End Function

    Private Function MessageToRow(ByVal zList As HashSet(Of String)) As List(Of Long)

        Dim sErr As String = ""
        Dim zOut As New List(Of Long)
        Dim FirstComma As Boolean = False
        Dim sQuery As New StringBuilder()

        If (zList.Count = 0) Then Return zOut

        sQuery.Append("SELECT rowid FROM spots WHERE msgid IN (")

        For Each sM As String In zList
            If FirstComma Then sQuery.Append(",") Else FirstComma = True
            sQuery.Append("'"c)
            sQuery.Append(sM.Replace("'"c, ""))
            sQuery.Append("'"c)
        Next

        sQuery.Append(") ORDER BY rowid DESC LIMIT " & zList.Count)

        Dim DR As DbDataReader = Db.ExecuteReader(sQuery.ToString(), sErr)
        If DR Is Nothing Then Throw New Exception("MessageToRow: " & sErr)

        While DR.Read
            zOut.Add(CLng(DR.Item(0)))
        End While

        DR.Close()

        Return zOut

    End Function

    Public Function Close() As Boolean

        Try

            If Not Db Is Nothing Then
                Db.Close()
                Db = Nothing
            End If

            Return True

        Catch ex As Exception

            Foutje(ex.Message)
            Return False

        End Try

    End Function

End Class
