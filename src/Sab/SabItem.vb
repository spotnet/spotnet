Imports System.ComponentModel
Imports System.Collections.ObjectModel

Imports Spotnet.Spotnet
Imports Spotlib

Friend Class SabItems
    Inherits ObservableCollection(Of SabItem)

    Public FD As New Dictionary(Of String, SabItem)
    Public Event ProgressChanged(ByVal lVal As Integer)

    Public Function Update(ByVal sID As String, ByVal bHistory As Boolean, ByVal Titel As String, ByVal Status As String, ByVal Perc As Integer, ByVal Omvang As String, ByVal TeGaan As String, ByVal zIndex As Integer, ByVal sLocation As String, ByVal zSpeed As Single) As Boolean

        Dim KL As SabItem
        Dim DoProgress As Boolean = False
        Dim ClearProgress As Boolean = False
        Dim DoMove As Boolean = False

        Dim IsDownloading As Boolean = False
        Dim WasDownloading As Boolean = False

        Dim IsCompleted As Boolean = False
        Dim WasCompleted As Boolean = False

        If FD.ContainsKey(sID.ToUpper) Then

            KL = FD.Item(sID.ToUpper)

            If KL.Perc <> Perc Then
                KL.Perc = Perc
                DoProgress = True
            End If

            WasCompleted = KL.IsCompleted
            WasDownloading = KL.IsDownloading

            If KL.Index <> zIndex Then KL.Index = zIndex
            If KL.History <> bHistory Then KL.History = bHistory
            If KL.Location <> sLocation Then KL.Location = sLocation
            If KL.RawStatus <> Status Then KL.Status = Status

            IsCompleted = KL.IsCompleted
            IsDownloading = KL.IsDownloading

            If Not IsDownloading Then zSpeed = 0

            If KL.RawTeGaan <> TeGaan Then KL.TeGaan = TeGaan
            If KL.RawSpeed <> zSpeed Then KL.RawSpeed = zSpeed

            If WasDownloading Then
                If Not IsDownloading Then
                    ClearProgress = True
                End If
            End If

            If IsDownloading Then
                If Not WasDownloading Then
                    DoProgress = True
                End If
            End If

            If IsCompleted Then
                If Not WasCompleted Then
                    Try
                        Dim Ref As MainWindow = CType(Application.Current.MainWindow, MainWindow)
                        Ref.DisplayTooltip(KL.Titel & " is compleet.")
                    Catch
                    End Try
                End If
            End If

        Else

            KL = New SabItem(sID, Titel, bHistory, Status, Perc, Omvang, TeGaan, zIndex, sLocation, zSpeed)

            If KL.IsHistory Or (KL.Index > MyBase.Count) Then
                MyBase.Add(KL)
            Else
                MyBase.Insert(KL.Index, KL)
            End If

            FD.Add(sID.ToUpper, KL)
            DoProgress = True

        End If

        If ClearProgress Then
            RaiseEvent ProgressChanged(-1)
        Else
            If DoProgress Then
                If KL.IsDownloading Then
                    If KL.Perc > 0 Then
                        RaiseEvent ProgressChanged(KL.Perc)
                    Else
                        RaiseEvent ProgressChanged(1)
                    End If

                End If
            End If
        End If

        Return True

    End Function

    Public Function RemoveID(ByVal sID As String) As Boolean

        Dim ClearProgress As Boolean = False

        If FD.ContainsKey(sID.ToUpper) Then
            Try
                If FD.Item(sID.ToUpper).IsDownloading Then ClearProgress = True
                MyBase.Remove(FD.Item(sID.ToUpper))
                FD.Remove(sID.ToUpper)
                If ClearProgress Then RaiseEvent ProgressChanged(-1)
                Return True
            Catch
            End Try
        End If

        Return False

    End Function

    Public Sub ClearAll()

        MyBase.Clear()
        FD.Clear()

    End Sub

End Class

Friend Class SabItem

    Implements INotifyPropertyChanged

    Private m_ID As String = ""
    Private m_Titel As String = ""
    Private m_Status As String = ""
    Private m_Perc As Integer = 0
    Private m_Index As Integer = 0
    Private m_Speed As Single = 0
    Private m_Omvang As String = ""
    Private m_Tegaan As String = ""
    Private m_History As Boolean
    Private m_Location As String = ""

    Public Event PropertyChanged As PropertyChangedEventHandler _
      Implements INotifyPropertyChanged.PropertyChanged

    Private Sub NotifyPropertyChanged(ByVal info As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(info))
    End Sub

    Public Sub New(ByVal ID As String, ByVal Titel As String, ByVal bHistory As Boolean, ByVal Status As String, ByVal Perc As Integer, ByVal Omvang As String, ByVal TeGaan As String, ByVal zIndex As Integer, ByVal Location As String, ByVal zSpeed As Single)

        m_ID = ID
        m_Perc = Perc
        m_Titel = Titel
        m_Omvang = Omvang
        m_Tegaan = TeGaan
        m_History = bHistory
        m_Index = zIndex
        m_Location = Location
        m_Speed = zSpeed
        m_Status = Status

    End Sub

    Public Property Perc As Integer

        Set(ByVal value As Integer)
            m_Perc = value
            NotifyPropertyChanged("Perc")
        End Set

        Get
            Return m_Perc
        End Get

    End Property

    Public Property Visibility As System.Windows.Visibility

        Set(ByVal value As System.Windows.Visibility)
            ''
        End Set

        Get
            If m_History Then
                Return System.Windows.Visibility.Hidden
            Else
                Return System.Windows.Visibility.Visible
            End If
        End Get

    End Property

    Public Property Opacity As Double

        Set(ByVal value As Double)
            ''
        End Set

        Get
            Return 1
        End Get

    End Property

    Public Property Speed As String

        Set(ByVal value As String)
            Tools.Foutje("Unsupported")
        End Set

        Get

            If IsDownloading Then
                If m_Speed > 0 Then Return Utils.ConvertSize(CLng(m_Speed * 1000)) & "/s"
            End If

            Return ""

        End Get

    End Property

    Public Property RawSpeed As Single

        Set(ByVal value As Single)
            m_Speed = value
            NotifyPropertyChanged("Speed")
        End Set

        Get
            Return m_Speed
        End Get

    End Property

    Public Property Index As Integer

        Set(ByVal value As Integer)
            m_Index = value
            NotifyPropertyChanged("Index")
        End Set

        Get
            Return m_Index
        End Get

    End Property

    Public Property ID As String

        Set(ByVal value As String)
            m_ID = value
            NotifyPropertyChanged("ID")
        End Set

        Get
            Return m_ID
        End Get

    End Property

    Public Property Location As String

        Set(ByVal value As String)
            m_Location = value
            NotifyPropertyChanged("Location")
        End Set

        Get
            Return m_Location
        End Get

    End Property

    Public Property Titel As String

        Set(ByVal value As String)
            m_Titel = value
            NotifyPropertyChanged("Titel")
        End Set

        Get
            Return m_Titel
        End Get

    End Property

    Public Property Omvang As String

        Set(ByVal value As String)
            m_Omvang = value
            NotifyPropertyChanged("Omvang")
        End Set

        Get
            Return m_Omvang
        End Get

    End Property

    Public Property TeGaan As String

        Set(ByVal value As String)
            m_Tegaan = value
            NotifyPropertyChanged("TeGaan")
        End Set

        Get
            If m_History Then
                Return ""
            Else
                If Val(m_Tegaan.Replace(":", "")) > 0 Then
                    Return m_Tegaan
                End If
                Return ""
            End If

        End Get

    End Property

    Public ReadOnly Property RawTeGaan As String

        Get
            Return m_Tegaan
        End Get

    End Property

    Public Property History As Boolean

        Set(ByVal value As Boolean)
            m_History = value
            NotifyPropertyChanged("History")
            NotifyPropertyChanged("Visibility")
        End Set

        Get
            Return m_History
        End Get

    End Property

    Public ReadOnly Property IsPaused As Boolean

        Get
            Return m_Status.ToLower = "paused"
        End Get

    End Property

    Public ReadOnly Property IsHistory As Boolean

        Get
            Return m_History
        End Get

    End Property

    Public ReadOnly Property IsDownloading As Boolean

        Get
            Return m_Status.ToLower = "downloading"
        End Get

    End Property

    Public ReadOnly Property IsCompleted As Boolean

        Get
            Return m_Status.ToLower = "completed"
        End Get

    End Property

    Public Property Status As String

        Set(ByVal value As String)

            m_Status = value
            NotifyPropertyChanged("Status")

        End Set

        Get

            Select Case m_Status.ToLower
                Case "completed" : Return "Compleet"
                Case "paused" : Return "Gepauzeerd"
                Case "failed" : Return "Mislukt"
                Case "queued" : Return "Wachten"
                Case "repairing" : Return "Repareren"
                Case "extracting" : Return "Uitpakken"
                Case "moving" : Return "Verplaatsen"
                Case "downloading" : Return "Downloaden"
                Case "running" : Return "Script uitvoeren"
                Case "fetching" : Return "Blokken ophalen"
                Case "quickcheck" : Return "Controleren"
                Case "verifying" : Return "Verifiëren"
                Case Else : Return m_Status
            End Select

        End Get

    End Property

    Public ReadOnly Property RawStatus As String

        Get
            Return m_Status
        End Get

    End Property

End Class
