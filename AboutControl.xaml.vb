Imports System.Windows.Media.Animation

Public Class AboutControl

    Friend Sub New()

        InitializeComponent()

    End Sub

    Private Sub AboutControl_Initialized(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Initialized

        Label2.Content = "v" & String.Format("{0}", My.Application.Info.Version.ToString).Substring(0, 5)

    End Sub

    Private Sub AboutControl_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        Dim storyboard As New Animation.Storyboard()
        Dim storyboard2 As New Animation.Storyboard()

        AddFade2(storyboard2, Grid1, 2)

        AddFade(storyboard, Image1, 2)
        AddFade(storyboard, Label1, 3)
        AddFade(storyboard, Label2, 3)
        AddFade(storyboard, Label3, 3)
        AddFade(storyboard, Label4, 3)
        AddFade(storyboard, Link, 3)

        ' Begin the storyboard
        storyboard.Begin(Me)
        storyboard2.Begin(Me)

    End Sub

    Private Sub AddFade(ByVal xBoard As Object, ByVal xIn As Object, ByVal sec As Integer)

        ' Create a DoubleAnimation to fade the not selected option control
        Dim animation As New DoubleAnimation()
        Dim duration As New TimeSpan(0, 0, sec)

        xIn.Opacity = 0
        xIn.Visibility = Visibility.Visible

        animation.From = 0.0
        animation.To = 1.0
        animation.Duration = New Duration(duration)

        ' Configure the animation to target de property Opacity
        Storyboard.SetTargetName(animation, CStr(xIn.Name))
        Storyboard.SetTargetProperty(animation, New PropertyPath(Control.OpacityProperty))
        ' Add the animation to the storyboard
        xBoard.Children.Add(animation)

    End Sub

    Private Sub AddFade2(ByVal xBoard As Object, ByVal xIn As Object, ByVal sec As Integer)

        ' Create a DoubleAnimation to fade the not selected option control
        Dim animation As New DoubleAnimation()
        Dim duration As New TimeSpan(0, 0, sec)

        xIn.Opacity = 1
        xIn.Visibility = Visibility.Visible

        animation.From = 1
        animation.To = 0.3
        animation.Duration = New Duration(duration)

        ' Configure the animation to target de property Opacity
        Storyboard.SetTargetName(animation, CStr(xIn.Name))
        Storyboard.SetTargetProperty(animation, New PropertyPath(Control.OpacityProperty))
        ' Add the animation to the storyboard
        xBoard.Children.Add(animation)

    End Sub

    Private Sub Linkje_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Linkje.Click

        LaunchBrowser("http://www.github.com/spotnet/spotnet/wiki")

    End Sub

End Class

