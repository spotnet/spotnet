Imports System.IO

Friend Class cTabs

    Public Tabs As New List(Of String)
    Private Const sFile As String = "\tabs.dat"

    Public Function SaveTabs(ByVal zTab As List(Of String)) As Boolean

        Try

            Dim SK As New StreamWriter(Tools.SettingsFolder() & sFile, False, System.Text.Encoding.UTF8)

            For Each xx As String In zTab
                If Len(xx) > 0 Then
                    SK.WriteLine(xx)
                End If
            Next

            SK.Close()
            Tabs = zTab

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Public Function ClearTabs() As Boolean

        Try

            System.IO.File.Delete(Tools.SettingsFolder() & sFile)
            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    Private Sub LT()

        If System.IO.File.Exists(Tools.SettingsFolder() & sFile) Then

            Dim SK As New StreamReader(Tools.SettingsFolder() & sFile, System.Text.Encoding.UTF8)

            Dim zS As String = ""

            Do While Not SK.EndOfStream
                zS = SK.ReadLine()
                If zS.Length > 0 Then Tabs.Add(zS)
            Loop

            SK.Close()

        End If

    End Sub

    Public Function LoadTabs() As Boolean

        Tabs = New List(Of String)

        Try


            LT()

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

End Class
