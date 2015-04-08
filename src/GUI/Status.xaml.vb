Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.ComponentModel
Imports System.Windows.Threading

Imports Spotlib
Imports Spotnet.Spotnet
Imports Spotbase.Spotbase

Public Class Status

    Friend TheSab As SabHelper

    Private ProgX As Long = -1
    Private ProgXM As String = ""

    Private TheGlass As New Glass
    Private MayUnload As Boolean = False

    Friend HeadersFile As String
    Friend CommentsFile As String

    Friend LastError As String = ""
    Friend HasError As Boolean = False

    Private WithEvents TheHeaders As Headers
    Private WithEvents TheComments As Comments

    Friend HeaderResults As rSaveSpots = Nothing
    Friend CommentResults As rSaveComments = Nothing

    Private WithEvents xSaveSpots As BackgroundWorker
    Private WithEvents xSaveComments As BackgroundWorker

    Public Event StatusChanged(ByVal sMsg As String, ByVal sVal As Long)

    Friend Sub New()

        InitializeComponent()

    End Sub

    Private Sub DoError(ByVal sLast As String)

        HasError = True
        LastError = sLast
        MayUnload = True

        Me.Close()

    End Sub

    Public Sub ProgressChanged(ByVal sMessage As String, ByVal sValue As Long)

        Try

            If Len(sMessage) > 0 Then
                If CType(Label1.Content, String) <> sMessage Then Label1.Content = sMessage
            End If

            If sValue = 0 Then

                If Not ProgressBar1.IsIndeterminate Then
                    ProgressBar1.IsIndeterminate = True
                End If

                If ProgressBar1.Value <> 0 Then ProgressBar1.Value = 0

            Else

                If sValue > 0 Then

                    If ProgressBar1.Value <> (sValue) Then ProgressBar1.Value = (sValue)
                    If ProgressBar1.IsIndeterminate Then ProgressBar1.IsIndeterminate = False
                Else

                    ProgressBar1.IsIndeterminate = False
                    ProgressBar1.Value = 0
                End If
            End If

            RaiseEvent StatusChanged("", sValue)

        Catch ex As Exception

            Static DidOnce As Boolean = False

            If Not DidOnce Then
                DidOnce = True
                Foutje("Progress_Changed:: " & ex.Message)
            End If

        End Try

    End Sub

    Private Function StartUpdate(ByRef zErr As String) As Boolean

        If Not TheHeaders Is Nothing Then
            zErr = ("Grabber <> Nothing")
            Return False
        End If

        If Not TheComments Is Nothing Then
            zErr = ("Comments <> Nothing")
            Return False
        End If

        Try

            Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)
            TheHeaders = Spots.FindSpots(HeaderPhuse, Ref.HeaderSettings(True, True))

            Return True

        Catch ex As Exception

            zErr = "StartUpdate: " & ex.Message
            Return False

        End Try

    End Function

    Private Sub Status_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing

        Try

            If Not MayUnload Then

                e.Cancel = True

            End If

            If Not e.Cancel Then

                If Not TheHeaders Is Nothing Then TheHeaders.Cancel()
                If Not TheComments Is Nothing Then TheComments.Cancel()

                Try
                    Me.Owner.Activate() 'WPF bug
                Catch
                End Try

            End If

        Catch ex As Exception
            Foutje("Status_Closing:: " & ex.Message)

        End Try

    End Sub

    Private Sub InitStatus()

        Static DidOnce As Boolean = False

        Try

            If Not DidOnce Then

                DidOnce = True

                ShowInit()

                If TheSab Is Nothing Then

                    Dim zErr As String = ""

                    If Not StartUpdate(zErr) Then

                        DoError(zErr)
                        Exit Sub

                    End If

                Else

                    Dim zErr As String = ""

                    If Not TheSab.StartSab(sModule.GetServer(ServerType.Download), False, zErr) Then
                        Foutje("Fout tijdens het starten van SABnzbd: " & zErr)
                        TheSab = Nothing
                    End If

                    MayUnload = True
                    Me.Close()
                    Exit Sub

                End If

            End If

        Catch ex As Exception

            Foutje("Status_Loaded:: " & ex.Message)

            MayUnload = True
            Me.Close()
            Exit Sub

        End Try

    End Sub

    Private Sub Status_Initialized(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Initialized

        Me.FontSize = My.Settings.FontSize + 1

        ShowInit()

    End Sub

    Private Sub ShowInit()

        If TheSab Is Nothing Then
            ProgressChanged("Verbinding maken...", 0)
        Else
            ProgressChanged("Downloads starten...", 0)
        End If

    End Sub

    Private Sub Status_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        ShowInit()
        Me.Dispatcher.BeginInvoke(New Action(AddressOf InitStatus), DispatcherPriority.Background)

    End Sub

    Private Sub TheHeaders_Completed(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs) Handles TheHeaders.Completed

        Try

            TheHeaders = Nothing

            Dim xReturn As SpotsCompletedEventArgs = CType(e, SpotsCompletedEventArgs)

            If xReturn.Cancelled Or (Not xReturn.Error Is Nothing) Then

                If xReturn.Error.Message <> CancelMSG Then

                    ProgressChanged(xReturn.Error.Message, -1)
                    DoError(xReturn.Error.Message)
                    Exit Sub

                End If

            Else

                If (xReturn.Spots.Count > 0) Or (xReturn.Deletes.Count > 0) Then

                    ProgressChanged("Spots opslaan...", 0)

                    CType(Application.Current.MainWindow, MainWindow).CloseDB()

                    xSaveSpots = New BackgroundWorker

                    xSaveSpots.WorkerReportsProgress = False
                    xSaveSpots.WorkerSupportsCancellation = False

                    xReturn.DbFile = HeadersFile

                    xSaveSpots.RunWorkerAsync(xReturn)

                    Exit Sub

                End If

                DoComments()

            End If

        Catch ex As Exception

            DoError("Grabber_Completed:: " & ex.Message)
            Exit Sub

        End Try

    End Sub

    Private Sub TheHeaders_ProgressChanged(ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles TheHeaders.ProgressChanged

        Dim zx As SpotnetProgressChangedEventArgs = CType(e, SpotnetProgressChangedEventArgs)
        ProgressChanged(zx.ProgressMessage, zx.ProgressPercentage)

    End Sub

    Private Sub Status_SourceInitialized(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SourceInitialized

        TheGlass.Init(Me)

    End Sub

    Private Sub DoComments()

        ProgressChanged("Reacties ophalen...", 0)

        Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)
        TheComments = Spots.FindComments(HeaderPhuse, Ref.CommentSettings(True, False))

    End Sub

    Private Sub xSaveSpots_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles xSaveSpots.DoWork

        Dim zError As String = ""

        Try

            Dim Saver As New Storage
            Dim SaveRes As rSaveSpots = Nothing
            Dim Ret As SpotsCompletedEventArgs = CType(e.Argument, SpotsCompletedEventArgs)

            e.Result = Nothing

            AddHandler Saver.ProgressChanged, AddressOf SaveProgressChanged

            If Saver.Connect(Ret.DbFile, zError) Then
                If Saver.InsertSpots(New Parameters(), Ret.Spots, Ret.Deletes, SaveRes, zError) Then
                    e.Result = SaveRes
                    Exit Sub
                End If
            End If

            LastError = zError

        Catch ex As Exception

            LastError = "SaveSpots: " & ex.Message

        End Try

    End Sub

    Private Sub xSaveSpots_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles xSaveSpots.RunWorkerCompleted

        Try

            xSaveSpots = Nothing

            Dim SaveRes As rSaveSpots = CType(e.Result, rSaveSpots)

            If SaveRes Is Nothing Then

                DoError(LastError)
                Exit Sub

            End If

            HeaderResults = SaveRes

            DoComments()

        Catch ex As Exception

            DoError("xSaveSpots_Completed:: " & ex.Message)
            Exit Sub

        End Try

    End Sub

    Private Sub TheComments_Completed(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs) Handles TheComments.Completed

        Try

            TheComments = Nothing

            Dim xReturn As CommentsCompletedEventArgs = CType(e, CommentsCompletedEventArgs)

            If xReturn.Cancelled Or (Not xReturn.Error Is Nothing) Then

                If xReturn.Error.Message <> CancelMSG Then

                    ProgressChanged(xReturn.Error.Message, -1)
                    DoError(xReturn.Error.Message)
                    Exit Sub

                End If

            Else

                If xReturn.Comments.Count > 0 Then

                    ProgressChanged("Reacties opslaan...", 0)

                    xSaveComments = New BackgroundWorker

                    xSaveComments.WorkerReportsProgress = False
                    xSaveComments.WorkerSupportsCancellation = False

                    xReturn.DbFile = CommentsFile

                    xSaveComments.RunWorkerAsync(xReturn)

                    Exit Sub

                End If

                ProgressChanged("Updaten voltooid", 100)

                MayUnload = True
                Me.Close()

            End If

        Catch ex As Exception

            DoError("Comments_Completed: " & ex.Message)
            Exit Sub

        End Try

    End Sub

    Private Sub TheComments_ProgressChanged(ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles TheComments.ProgressChanged

        Dim zx As SpotnetProgressChangedEventArgs = CType(e, SpotnetProgressChangedEventArgs)
        ProgressChanged(zx.ProgressMessage, zx.ProgressPercentage)

    End Sub

    Private Sub xSaveComments_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles xSaveComments.DoWork

        Dim zError As String = ""

        Try

            Dim Saver As New Storage
            Dim SaveRes As rSaveComments = Nothing
            Dim Ret As CommentsCompletedEventArgs = CType(e.Argument, CommentsCompletedEventArgs)

            e.Result = Nothing

            AddHandler Saver.ProgressChanged, AddressOf SaveProgressChanged

            If Saver.Connect(Ret.DbFile, zError) Then
                If Saver.AddComments(New Parameters(), Ret.Comments, SaveRes, zError) Then
                    e.Result = SaveRes
                    Exit Sub
                End If
            End If

            LastError = zError

        Catch ex As Exception

            LastError = "SaveComments: " & ex.Message

        End Try

    End Sub

    Private Sub xSaveComments_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles xSaveComments.RunWorkerCompleted

        Try

            xSaveComments = Nothing

            Dim SaveRes As rSaveComments = CType(e.Result, rSaveComments)

            If SaveRes Is Nothing Then
                DoError(LastError)
                Exit Sub
            End If

            CommentResults = SaveRes
            ProgressChanged("Updaten voltooid", 100)

            MayUnload = True
            Me.Close()

        Catch ex As Exception

            DoError("xSaveSpots_Completed:: " & ex.Message)
            Exit Sub

        End Try

    End Sub

    Private Sub SetP()

        ProgressChanged(ProgXM, ProgX)

    End Sub

    Public Sub SaveProgressChanged(ByVal sMsg As String, ByVal lVal As Integer)

        ProgX = lVal
        ProgXM = sMsg

        Me.Dispatcher.BeginInvoke(New Action(AddressOf Me.SetP), DispatcherPriority.Normal)

    End Sub

End Class
