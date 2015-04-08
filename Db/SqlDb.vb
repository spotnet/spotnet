Imports System.Text

Imports System.Data.Common
Imports System.Data.SQLite

Friend Class SqlDB

    Private CNN As DbConnection
    Private ThreadID As Integer = -1
    Private Fact As New System.Data.SQLite.SQLiteFactory

    Public Function Connect(ByVal dbFile As String, ByVal bReadOnly As Boolean) As Boolean

        Close()

        If Len(dbFile) = 0 Then
            Throw New Exception("No database specified!")
        End If

        CNN = Fact.CreateConnection
        CNN.ConnectionString = "Data Source=" & dbFile & ";" & sIIF(bReadOnly, "Read Only=True;", "")
        CNN.Open()

        ThreadID = System.Threading.Thread.CurrentThread.ManagedThreadId()

        Return CNN.State = ConnectionState.Open

    End Function

    Public ReadOnly Property Connected As Boolean

        Get
            Return (Not (CNN Is Nothing))
        End Get

    End Property

    Private Function CheckThread() As Boolean

        If (ThreadID <> -1) And (ThreadID <> System.Threading.Thread.CurrentThread.ManagedThreadId()) Then
            Foutje("Wrong thread")
            Return False
        End If

        Return True

    End Function

    Public Function ExecuteNonQuery(ByVal sQuery As String, ByRef sErr As String) As Integer

        Try

            If sQuery.Length = 0 Then
                sErr = "No query"
                Return -2
            End If

            If Not CheckThread() Then
                sErr = "Wrong thread"
                Return -2
            End If

            Dim dbCmd As DbCommand = CNN.CreateCommand
            dbCmd.CommandText = sQuery
            Return (CType(dbCmd.ExecuteNonQuery(), Integer))

        Catch ex As Exception
            sErr = ex.Message
        End Try

        Return -2

    End Function

    Public Function ExecuteReader(ByVal sQuery As String, ByRef sErr As String) As DbDataReader

        Try

            If sQuery.Length = 0 Then
                sErr = "No query"
                Return Nothing
            End If

            If Not CheckThread() Then
                sErr = "Wrong thread"
                Return Nothing
            End If

            Dim dbCmd As DbCommand = CNN.CreateCommand
            dbCmd.CommandText = sQuery
            Return dbCmd.ExecuteReader()

        Catch ex As Exception

            sErr = ex.Message
            Return Nothing

        End Try

    End Function

    Public Function ExecuteScalar(ByVal sQuery As String, ByRef sErr As String) As Long

        Try

            If sQuery.Length = 0 Then
                sErr = "No query"
                Return -1
            End If

            If Not CheckThread() Then
                sErr = "Wrong thread"
                Return -1
            End If

            Dim dbCmd As DbCommand = CNN.CreateCommand

            dbCmd.CommandText = sQuery

            Return CType(dbCmd.ExecuteScalar(), Long)

        Catch ex As Exception

            sErr = ex.Message
            Return -1

        End Try

    End Function

    Public Function ExecuteCommand(ByVal sQuery As String) As String

        Dim sErr As String = ""
        Dim sRet As String = ""
        Dim bFirstField As Boolean = True
        Dim DR As DbDataReader = Nothing
        Dim zRet As New StringBuilder()

        Try

            DR = ExecuteReader(sQuery, sErr)

            If DR Is Nothing Then
                Throw New Exception(sErr)
            End If

            While DR.Read
                For i As Integer = 0 To DR.FieldCount - 1
                    If bFirstField Then
                        bFirstField = False
                    Else
                        zRet.Append(" "c)
                    End If
                    If Not IsDBNull(DR.Item(i)) Then
                        zRet.Append(CStr(DR.Item(i)))
                    End If
                Next
                zRet.Append(vbCrLf)
                bFirstField = True
            End While

            DR.Close()
            DR = Nothing

            Return zRet.ToString()

        Catch ex As Exception

            Try
                DR.Close()
            Catch
            End Try

            Throw New Exception("ExecuteCommand: " & ex.Message)

        End Try

    End Function

    Public Function BeginTransaction() As DbTransaction

        If Not CheckThread() Then
            Throw New Exception("Wrong thread")
        End If

        Return CNN.BeginTransaction()

    End Function

    Public Function CreateCommand() As DbCommand

        If Not CheckThread() Then
            Throw New Exception("Wrong thread")
        End If

        Return CNN.CreateCommand()

    End Function

    Public Sub Close()

        If Not CNN Is Nothing Then

            Try

                CNN.Close()
                CNN.Dispose()

            Catch
            End Try

            CNN = Nothing

        End If

    End Sub

    Protected Overrides Sub Finalize()

        Close()
        MyBase.Finalize()

    End Sub

End Class