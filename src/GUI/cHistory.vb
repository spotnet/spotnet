Imports System.IO

Friend Class cHistory

    Public History As New List(Of String)
    Private Const sFile As String = "\history.dat"

    Public Function SaveHistory(ByVal zHis As String) As Boolean

        Try

            If zHis.Length > 0 Then
                If Not History.Contains(zHis) Then

                    Dim SK As New StreamWriter(Tools.SettingsFolder() & sFile, True, System.Text.Encoding.UTF8)
                    SK.WriteLine(zHis)
                    SK.Close()

                    History.Add(zHis)

                End If
            End If


            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Public Function ClearHistory() As Boolean

        Try

            System.IO.File.Delete(Tools.SettingsFolder() & sFile)
            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Private Sub LH()

        If System.IO.File.Exists(Tools.SettingsFolder() & sFile) Then

            Dim SK As New StreamReader(Tools.SettingsFolder() & sFile, System.Text.Encoding.UTF8)

            Dim zS As String = ""

            Do While Not SK.EndOfStream
                zS = SK.ReadLine()
                If zS.Length > 0 Then History.Add(zS)
            Loop

            SK.Close()

        End If

    End Sub

    Public Function LoadHistory() As Boolean

        History = New List(Of String)

        Try

            LH()

            If History.Count > 1000 Then

                Dim SK As New StreamWriter(Tools.SettingsFolder() & sFile, False, System.Text.Encoding.UTF8)

                Dim cI As Long = 1

                For Each Xl As String In History
                    cI += 1
                    If cI > 500 Then SK.WriteLine(Xl)
                Next

                SK.Close()

                History = New List(Of String)

                LH()

            End If

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

End Class
