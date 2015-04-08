Imports System.IO
Imports System.Windows.Media.Imaging

Imports Spotlib

Friend Class SpotInfo

    Public Spot As Spotlib.SpotEx
    Public TabLoaded As Boolean

End Class

Friend Class UrlInfo

    Public URL As String
    Public Title As String
    Public TabLoaded As Boolean

End Class

Public Class ListItem

    Public Key As String
    Public Name As String

    Public Sub New(ByVal sKey As String, ByVal sName As String)

        Key = sKey
        Name = sName

    End Sub

End Class

Friend Class SabSlots

    Public DataSlots As List(Of String())
    Public HistorySlots As List(Of String())

End Class

Friend Class Tools

    Public Shared Function SettingsFolder() As String

        Static TheResult As String = ""
        Static DidOnce As Boolean = False

        If Not DidOnce Then

            DidOnce = True

            Dim AddFolder As String = Spotz.Spotname

            If Utils.FileExists(AppPath() & "\settings.xml") Then

                TheResult = AppPath()

            Else

                Dim Myfolder As String = ""

                Myfolder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData)

                If Not Utils.DirectoryExists(Myfolder & "\" & AddFolder) Then
                    If Directory.CreateDirectory(Myfolder & "\" & AddFolder) Is Nothing Then
                        Foutje("Kan '" & Myfolder & "\" & AddFolder & "' niet aanmaken!")
                    End If
                End If

                TheResult = Myfolder & "\" & AddFolder

            End If
        End If

        Return TheResult

    End Function

    Public Shared Sub CS(Optional ByVal lMil As Integer = 1000)

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

    Public Shared Sub Foutje(ByVal sMsg As String, Optional ByVal sCaption As String = "Oeps")

        CS(0)
        Debug.Print(sMsg)

        Try
            MsgBox(sMsg, CType(MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, MsgBoxStyle), sCaption)
        Catch
        End Try

    End Sub

    Public Shared Function AppPath() As String

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

    Public Shared Function DownDir() As String

        If Len(Trim(My.Settings.DownloadFolder)) > 0 Then
            If Utils.DirectoryExists(My.Settings.DownloadFolder) Then
                Return My.Settings.DownloadFolder
            End If
        End If

        If (Environment.OSVersion.Version.Major >= 6) Then
            Return Utils.GetFolder(New Guid("374DE290-123F-4565-9164-39C4925E467B"))
        Else
            Return Environment.ExpandEnvironmentVariables("%USERPROFILE%") & "\Downloads"
        End If

    End Function

    Friend Shared Function CreateHeader(ByVal sTitle As String, ByVal sIcon As String, Optional ByVal ISX As ImageSource = Nothing) As StackPanel

        Dim KL As New StackPanel
        KL.Orientation = Orientation.Horizontal

        Dim kli As New Image
        If ISX Is Nothing Then kli.Source = Common.GetImage(sIcon) Else kli.Source = ISX
        kli.VerticalAlignment = Windows.VerticalAlignment.Bottom

        Dim klt As New TextBlock
        klt.Text = sTitle.Trim & " "
        klt.Margin = New Thickness(4, 0, 0, 0)
        klt.VerticalAlignment = Windows.VerticalAlignment.Center

        KL.Children.Add(kli)
        KL.Children.Add(klt)

        Return KL

    End Function

    Friend Shared Function GetHeader(ByVal Zt As Object) As String

        Dim kd As StackPanel = CType(Zt, StackPanel)
        Dim pd As TextBlock = CType(kd.Children(1), TextBlock)
        Return pd.Text

    End Function

    Friend Shared Function GetHeaderIcon(ByVal Zt As Object) As ImageSource

        Dim kd As StackPanel = CType(Zt, StackPanel)
        Dim pd As Image = CType(kd.Children(0), Image)
        Return pd.Source

    End Function

    Friend Shared Function SpotTheme() As String

        Dim sFile As String = "spot"
        Dim objReader As StreamWriter

        sFile = SettingsFolder() & "\Theme\" & sFile & ".htm"

        If Not Utils.FileExists(sFile) Then
            Try

                If Not Utils.DirectoryExists(SettingsFolder() & "\Theme") Then
                    Directory.CreateDirectory(SettingsFolder() & "\Theme")
                End If

                objReader = New StreamWriter(sFile, False, Utils.LatinEnc)
                objReader.Write(My.Resources.spot.ToString())
                objReader.Close()

            Catch Ex As Exception
            End Try
        End If

        Return sFile

    End Function

    Friend Shared Function CommentTheme() As String

        Dim sFile As String = "comment"
        Dim objReader As StreamWriter

        sFile = SettingsFolder() & "\Theme\" & sFile & ".htm"

        If Not Utils.FileExists(sFile) Then

            Try

                If Not Utils.DirectoryExists(SettingsFolder() & "\Theme") Then
                    Directory.CreateDirectory(SettingsFolder() & "\Theme")
                End If

                objReader = New StreamWriter(sFile, False, Utils.LatinEnc)
                objReader.Write(My.Resources.comment.ToString())
                objReader.Close()

            Catch Ex As Exception
            End Try
        End If

        Return sFile

    End Function

    Friend Shared Sub ShowOnce(ByVal sMsg As String, ByVal sTitle As String)

        Static DidList As New HashSet(Of String)

        If DidList.Add(sMsg) Then
            MsgBox(sMsg, MsgBoxStyle.Information, sTitle)
        End If

    End Sub

    Public Shared Function GetAvatar() As Byte()

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

End Class

Public Class WorkParams

    Implements Spotlib.iWorkParams

    Public Sub Save() Implements iWorkParams.Save

        My.Settings.Save()

    End Sub

    Public Function DatabaseCache() As Integer Implements iWorkParams.DatabaseCache

        Return My.Settings.DatabaseCache

    End Function

    Public Function DatabaseCount() As Long Implements iWorkParams.DatabaseCount

        Return My.Settings.DatabaseCount

    End Function

    Public Function DatabaseFilter() As Long Implements iWorkParams.DatabaseFilter

        Return My.Settings.DatabaseFilter

    End Function

    Public Function DatabaseMax() As Long Implements iWorkParams.DatabaseMax

        Return My.Settings.DatabaseMax

    End Function

    Public Function MaxResults() As Integer Implements iWorkParams.MaxResults

        Return My.Settings.MaxResults

    End Function

    Public Sub SetDatabaseCache(vRes As Integer) Implements iWorkParams.SetDatabaseCache

        My.Settings.DatabaseCache = vRes

    End Sub

    Public Sub SetDatabaseCount(vRes As Long) Implements iWorkParams.SetDatabaseCount

        My.Settings.DatabaseCount = vRes

    End Sub

    Public Sub SetDatabaseFilter(vRes As Long) Implements iWorkParams.SetDatabaseFilter

        My.Settings.DatabaseFilter = vRes

    End Sub

    Public Sub SetDatabaseMax(vRes As Long) Implements iWorkParams.SetDatabaseMax

        My.Settings.DatabaseMax = vRes

    End Sub

    Public Sub SetMaxResults(vRes As Integer) Implements iWorkParams.SetMaxResults

        My.Settings.MaxResults = vRes

    End Sub

End Class