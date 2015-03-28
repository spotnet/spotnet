Imports System.Text
Imports System.Data.SQLite
Imports System.Data.Common
Imports DataVirtualization

Friend Class SpotProvider

    Implements DataVirtualization.IVirtualListLoader(Of SpotRow)

    Private Db As SqlDB
    Private CacheFile As String
    Private CacheQueryCount As Integer = -1
    Private CacheQueryCounts As New Dictionary(Of String, Integer)

    Private _RowNew As Long = -1
    Private _Corrupted As Boolean = False

    Private Const DefCols2 As String = "cats TEXT, sender TEXT, tag TEXT, subject TEXT"
    Private Const DefCols As String = "key INT, cat INT, subcat INT, extcat INT, date INT, filesize INTEGER, cats TEXT, sender TEXT, tag TEXT, subject TEXT, msgid TEXT, modulus TEXT"

    Private _RowFilter As String = DefaultFilter
    Private _QueryName As String = _QueryDefaultName
    Private _QueryDefaultName As String = "Overzicht"

    Public Property QueryName As String

        Get
            If Len(_QueryName) = 0 Or (_QueryName = sModule.DefaultFilter) Then
                Return _QueryDefaultName
            Else
                Return _QueryName
            End If
        End Get

        Set(ByVal value As String)
            If Len(value) = 0 Or (_QueryName = sModule.DefaultFilter) Then
                _QueryName = _QueryDefaultName
            Else
                _QueryName = value
            End If
            _QueryName = value
        End Set

    End Property

    Public Property SortCol As String

        Get
            Return My.Settings.SortColumn
        End Get

        Set(ByVal value As String)
            If My.Settings.SortColumn <> value Then
                My.Settings.SortColumn = value
                My.Settings.Save()
            End If
        End Set

    End Property

    Public Property SortOrder As String

        Get
            Return My.Settings.SortDirection
        End Get

        Set(ByVal value As String)
            If My.Settings.SortDirection <> value Then
                My.Settings.SortDirection = value
                My.Settings.Save()
            End If
        End Set

    End Property

    Private Function GetString(ByVal obj As Object) As String

        If IsDBNull(obj) Then Return ""
        Return CStr(obj)

    End Function

    Friend Sub ResetCount()

        CacheQueryCount = -1

    End Sub

    Private Sub ResetCache()

        ResetCount()
        CacheQueryCounts = New Dictionary(Of String, Integer)

    End Sub

    Public Property RowFilter() As String

        Get
            If _RowFilter Is Nothing Then Return ""
            Return _RowFilter
        End Get

        Set(ByVal value As String)
            If _RowFilter <> value Then
                _RowFilter = value
            End If
        End Set

    End Property

    Friend ReadOnly Property Connected As Boolean

        Get
            If Db Is Nothing Then Return False
            Return Db.Connected
        End Get

    End Property

    Friend ReadOnly Property Filename As String

        Get
            Return CacheFile
        End Get

    End Property

    Public ReadOnly Property CanSort As Boolean Implements DataVirtualization.IVirtualListLoader(Of SpotRow).CanSort

        Get
            Return True
        End Get

    End Property

    Private Function fDate() As Integer

        Static xDate As Integer = 0

        If xDate = 0 Then
            Dim Span As TimeSpan = (DateTime.UtcNow - EPOCH)
            xDate = CInt(Span.TotalSeconds)
        End If

        Return xDate

    End Function

    Public Function LoadRange(ByVal startIndex As Integer, ByVal Count As Integer, ByVal sortDescriptions As System.ComponentModel.SortDescriptionCollection, ByRef overallCount As Integer) As System.Collections.Generic.IList(Of SpotRow) Implements DataVirtualization.IVirtualListLoader(Of SpotRow).LoadRange

        Dim sErr As String = ""
        Dim sOffset As String = ""
        Dim lPosIndex As Integer = 0
        Dim DR As DbDataReader = Nothing
        Dim CountQuery As String = Nothing

        Dim Perf As New Stopwatch

        If Not Connected Then
            overallCount = 0
            Return New List(Of SpotRow)
        End If

        Try

            lPosIndex = 1
            Perf.Start()

            Debug.Print("LoadRange: " & CStr(startIndex))

            Mouse.OverrideCursor = Cursors.Wait

            Dim MaxResults As Integer = CInt(My.Settings.MaxResults)
            If MaxResults <= 0 Then MaxResults = -1 ' Oude setting

            If CacheQueryCount < 1 Then
                CacheQueryCount = -1
                If ((_RowFilter = DefaultFilter) Or (Len(_RowFilter) = 0)) Then
                    CacheQueryCount = CInt(My.Settings.DatabaseFilter)
                End If
            End If

            lPosIndex = 2

            If MaxResults > 1 Then
                If (startIndex) > MaxResults Then
                    overallCount = 0
                    Return New List(Of SpotRow)
                Else
                    If (Count) > MaxResults Then
                        Count = MaxResults
                    End If
                    If (startIndex + Count) > MaxResults Then
                        Count = MaxResults - startIndex
                    End If
                End If
            End If

            sOffset = " LIMIT " & Count & " "
            If startIndex > 0 Then sOffset += "OFFSET " & startIndex & " "

            lPosIndex = 3

            Dim sQ As String

            If Not IsSearchQuery(_RowFilter) Then
                sQ = CreateQuery(MaxResults, sOffset, CountQuery)
            Else
                sQ = CreateSearchQuery(MaxResults, sOffset, CountQuery)
            End If

            lPosIndex = 4

            DR = Db.ExecuteReader(sQ, sErr)

            lPosIndex = 5

            If DR Is Nothing Then
                Throw New Exception("Datareader not available: " & sErr)
            End If

            lPosIndex = 6

            Dim SR() As SpotRow = ReadRows(DR, Count, CountQuery)

            lPosIndex = 7

            If (CacheQueryCount > MaxResults) And (MaxResults > 0) Then
                CacheQueryCount = MaxResults
            End If

            overallCount = CacheQueryCount

            lPosIndex = 9
            Perf.Stop()

            Debug.Print("SQL Time: " & Perf.ElapsedMilliseconds & " ms")

            Return SR

        Catch ex As Exception

            Try
                DR.Close()
            Catch
            End Try

            Mouse.OverrideCursor = Nothing

            Foutje("LoadRange: #" & lPosIndex & " - " & ex.Message)

            If TypeOf (ex) Is SQLiteException Then

                Dim Sq As SQLiteException = CType(ex, SQLiteException)

                Select Case Sq.ErrorCode
                    Case SQLiteErrorCode.Corrupt, SQLiteErrorCode.NotADatabase
                        _Corrupted = True
                End Select

            End If

            ResetCache()
            overallCount = 0

            Return Nothing

        End Try

    End Function

    Private Function CreateSearchQuery(ByVal MaxResults As Integer, ByVal sOffset As String, ByRef CountQuery As String) As String

        Dim sQ As String = ""
        Dim RowSort As String = ""

        If Len(_RowFilter.Trim) > 0 Then
            sQ = " WHERE " & _RowFilter
        End If

        If sQ.Contains("[SN:DATE]") Then
            sQ = Strings.Replace(sQ, "[SN:DATE]", CStr(fDate()), 1, , CompareMethod.Binary)
        End If

        If sQ.Contains("[SN:NEW]") Then
            If _RowNew > 1 Then
                sQ = Strings.Replace(sQ, "[SN:NEW]", Convert.ToString(_RowNew), 1, , CompareMethod.Binary)
            Else
                sQ = Strings.Replace(sQ, "[SN:NEW]", Convert.ToString(My.Settings.DatabaseMax + 1), 1, , CompareMethod.Binary)
            End If
        End If

        Dim SelectID As String = "docid"
        Dim SelectTable As String = "search"

        Dim TotalCols As String = "subcat,extcat,date,filesize,subject,sender,tag,modulus"
        Dim SortCol As String = My.Settings.SortColumn.Trim.ToLower
        Dim ExternalSort As Boolean = False

        Select Case SortCol
            Case "subcat", "extcat", "date", "filesize"
                ExternalSort = True
        End Select

        RowSort = " ORDER BY " & SortCol & " " & SortOrder & " "
        If SortCol <> "rowid" Then RowSort += ", rowid DESC "

        CountQuery = "SELECT COUNT(*) FROM "

        If (MaxResults < 1) Then
            CountQuery += SelectTable & sQ
        Else
            CountQuery += "(SELECT null FROM " & SelectTable & sQ & " LIMIT " & MaxResults & ")"
        End If

        Dim sPrefix As String = "SELECT rowid," & TotalCols & " FROM spots WHERE rowid IN "

        If Not ExternalSort Then
            If (MaxResults < 1) Or (SortCol = "rowid" And SortOrder.ToLower = "desc") Then
                sQ = sPrefix & "(SELECT " & SelectID & " FROM " & SelectTable & sQ & RowSort & sOffset & ")" & RowSort
            Else
                sQ = sPrefix & "(SELECT " & SelectID & " FROM " & SelectTable & sQ & " ORDER BY " & SelectID & " DESC" & " LIMIT " & MaxResults & ") " & RowSort & sOffset
            End If
        Else
            If (MaxResults < 1) Then
                sQ = sPrefix & "(SELECT " & SelectID & " FROM " & SelectTable & sQ & ") " & RowSort & sOffset
            Else
                sQ = sPrefix & "(SELECT rowid FROM spots WHERE rowid IN (SELECT " & SelectID & " FROM " & SelectTable & sQ & ") ORDER BY rowid DESC LIMIT " & MaxResults & ")" & RowSort & sOffset
            End If
        End If

        Debug.Print("SQL Query: " & sQ)

        Return sQ

    End Function

    Private Function CreateQuery(ByVal MaxResults As Integer, ByVal sOffset As String, ByRef CountQuery As String) As String

        Dim sQ As String = ""
        Dim RowSort As String = ""

        If Len(_RowFilter.Trim) > 0 Then
            sQ = " WHERE " & _RowFilter
        End If

        If sQ.Contains("[SN:DATE]") Then sQ = Strings.Replace(sQ, "[SN:DATE]", CStr(fDate()), 1, , CompareMethod.Binary)

        Dim SelectID As String = "rowid"
        Dim SelectTable As String = "spots"

        Dim TotalCols As String = "subcat,extcat,date,filesize,subject,sender,tag,modulus"
        Dim SortCol As String = My.Settings.SortColumn.Trim.ToLower

        RowSort = " ORDER BY " & SortCol & " " & SortOrder & " "
        If SortCol <> "rowid" Then RowSort += ", rowid DESC "

        CountQuery = "SELECT COUNT(*) FROM "

        If (MaxResults < 1) Then
            CountQuery += SelectTable & sQ
        Else
            CountQuery += "(SELECT null FROM " & SelectTable & sQ & " LIMIT " & MaxResults & ")"
        End If

        Dim sPrefix As String = "SELECT " & SelectID & "," & TotalCols & " FROM "

        If (MaxResults < 1) Or (SortCol = "rowid" And SortOrder.ToLower = "desc") Then
            sQ = sPrefix & SelectTable & sQ & RowSort & sOffset
        Else
            sQ = sPrefix & SelectTable & " WHERE " & SelectID & " IN (SELECT " & SelectID & " FROM " & SelectTable & sQ & " ORDER BY " & SelectID & " DESC" & " LIMIT " & MaxResults & ") " & RowSort & sOffset
        End If

        Debug.Print("SQL Query: " & sQ)

        Return sQ

    End Function

    Private Function ReadRows(ByVal DR As DbDataReader, ByVal lCount As Integer, ByVal CountQuery As String) As SpotRow()

        Dim lCnt As Integer = 0
        Dim sErr As String = ""
        Dim SR(lCount - 1) As SpotRow

        With DR
            While .Read

                Try

                    Dim SP As New SpotRowChild

                    SP.ID = CLng(.Item(0))
                    SP.SubCat = CInt(.Item(1))
                    SP.ExtCat = CInt(.Item(2))
                    SP.Stamp = CInt(.Item(3))
                    SP.Filesize = CLng(.Item(4))

                    SP.Title = GetString(.Item(5))
                    SP.Poster = GetString(.Item(6))
                    SP.Tag = GetString(.Item(7))
                    SP.Modulus = GetString(.Item(8))

                    Dim Sb As New iSpotRow(SP)

                    SR(lCnt) = Sb

                    lCnt += 1

                    If lCnt = lCount Then

                        If CacheQueryCount < 0 Then

                            DR.Close()

                            If CacheQueryCounts.ContainsKey(CountQuery) Then

                                CacheQueryCount = CacheQueryCounts(CountQuery)

                            Else

                                CacheQueryCount = CInt(Db.ExecuteScalar(CountQuery, sErr))
                                If CacheQueryCount < 0 Then Throw New Exception(sErr)

                                CacheQueryCounts.Add(CountQuery, CacheQueryCount)

                                Debug.Print("SQL Count: " & CacheQueryCount & " - " & CountQuery)

                            End If

                        End If

                        Exit While

                    End If

                Catch ex As Exception

                    Throw New Exception("Reader Error: " & ex.Message)

                End Try

            End While
        End With

        DR.Close()

        If lCnt > CacheQueryCount Then
            CacheQueryCount = lCnt
        End If

        Do While lCnt < lCount
            SR(lCnt) = New iSpotRow(Nothing)
            lCnt += 1
        Loop

        Return SR

    End Function

    Friend ReadOnly Property QueryCount As Long

        Get
            If CacheQueryCount > 0 Then Return CacheQueryCount Else Return 0
        End Get

    End Property

    Friend Function GetMessageID(ByVal sTable As String, ByVal ID As Long) As String

        Dim sErr As String = ""
        Dim sRet As String = ""

        Try

            Return Db.ExecuteCommand("SELECT msgid FROM " & sTable & " WHERE rowid = " & ID).Replace(vbCrLf, "").Trim

        Catch ex As Exception

            Throw New Exception("GetMessageID: " & ex.Message)

        End Try

    End Function

    Private Sub CreateTables()

        If Not Db.ExecuteNonQuery("PRAGMA page_size = 4096;", "") = 0 Then Throw New Exception("PRAGMA page_size")
        If Not Db.ExecuteNonQuery("PRAGMA synchronous = OFF", "") = 0 Then Throw New Exception("PRAGMA synchronous")
        If Not Db.ExecuteNonQuery("PRAGMA journal_mode = MEMORY", "") = 0 Then Throw New Exception("PRAGMA journal_mode")
        If Not Db.ExecuteNonQuery("PRAGMA temp_store = MEMORY;", "") = 0 Then Throw New Exception("PRAGMA temp_store")
        If Not Db.ExecuteNonQuery("PRAGMA locking_mode = EXCLUSIVE", "") = 0 Then Throw New Exception("PRAGMA locking_mode")
        If Not Db.ExecuteNonQuery("PRAGMA cache_size = " & CStr(My.Settings.DatabaseCache), "") = 0 Then Throw New Exception("PRAGMA cache_size")

        Dim dbTrans As DbTransaction = Db.BeginTransaction

        If Not Db.ExecuteNonQuery("CREATE TABLE spots(rowid INTEGER PRIMARY KEY, " & DefCols & " )", "") = 0 Then Throw New Exception("CREATE TABLE spots")
        If Not Db.ExecuteNonQuery("CREATE VIRTUAL TABLE search USING fts4(content=" & Chr(34) & "spots" & Chr(34) & "," & DefCols2 & ",order=desc,matchinfo=fts3)", "") = 0 Then Throw New Exception("CREATE TABLE search")

        dbTrans.Commit()
        dbTrans.Dispose()

    End Sub

    Public Function Connect(ByVal dbFile As String, ByRef sError As String) As Boolean

        Dim sErr As String = ""
        Dim bResetTotal As Boolean = False
        Dim bNew As Boolean = Not FileExists(dbFile)

        _Corrupted = False

        Try

            Close()
            Db = New SqlDB

            If Len(CacheFile) > 0 Then
                If CacheFile <> dbFile Then bResetTotal = True
            End If

            CacheFile = dbFile

            If Not bNew Then
                If FileLen(CacheFile) < 1 Then bNew = True
            End If

            If bNew Then

                bResetTotal = True

                If Not Db.Connect(dbFile, False) Then Throw New Exception("iConnect")

                CreateTables()

                Close()

                If (FileLen(CacheFile) < 1) Then Throw New Exception(sErr)

                Db = New SqlDB

            End If

            If Not Db.Connect(dbFile, True) Then Throw New Exception("iConnect")

            If (My.Settings.DatabaseMax < 1) Then bResetTotal = True
            If (My.Settings.DatabaseCount < 1) Then bResetTotal = True
            If (My.Settings.DatabaseFilter < 1) Then bResetTotal = True

            Dim cMax As Long = Db.ExecuteScalar("SELECT MAX(rowid) FROM spots", "")

            If cMax < 1 Then bResetTotal = True
            If cMax <> My.Settings.DatabaseMax Then bResetTotal = True

            If bResetTotal Then

                My.Settings.DatabaseMax = cMax
                My.Settings.DatabaseCount = Db.ExecuteScalar("SELECT COUNT(*) FROM spots", "")
                My.Settings.DatabaseFilter = My.Settings.DatabaseCount - Math.Abs(Db.ExecuteScalar("SELECT COUNT(*) FROM spots WHERE cat=9", ""))
                My.Settings.Save()

            End If

            If Not Db.ExecuteNonQuery("PRAGMA temp_store = MEMORY;", "") = 0 Then Throw New Exception("PRAGMA temp_store")
            If Not Db.ExecuteNonQuery("PRAGMA cache_size = " & CStr(My.Settings.DatabaseCache), "") = 0 Then Throw New Exception("PRAGMA cache_size")

            Return True

        Catch ex As AccessViolationException

            If Len(sErr) > 0 Then
                sError = sErr & vbCrLf
            End If

            sError += ex.Message
            Return False

        Catch ex As Exception

            If Len(sErr) > 0 Then
                sError = sErr & vbCrLf
            End If

            If TypeOf (ex) Is SQLiteException Then

                Dim Sq As SQLiteException = CType(ex, SQLiteException)

                Select Case Sq.ErrorCode
                    Case SQLiteErrorCode.Corrupt, SQLiteErrorCode.NotADatabase
                        _Corrupted = True
                End Select

            End If

            sError += ex.Message
            Return False

        End Try

    End Function

    Public ReadOnly Property Corrupted As Boolean

        Get
            Return _Corrupted
        End Get

    End Property

    Public Property RowNew As Long

        Get
            Return _RowNew
        End Get

        Set(ByVal value As Long)
            _RowNew = value
        End Set

    End Property

    Friend Function Close() As Boolean

        _RowNew = -1

        ResetCache()

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

    Public ReadOnly Property LastPosition(ByVal sTable As String) As Long

        Get

            Return sModule.LastPosition(Db, sTable)

        End Get

    End Property

    Protected Overrides Sub Finalize()

        Close()

        MyBase.Finalize()

    End Sub

End Class