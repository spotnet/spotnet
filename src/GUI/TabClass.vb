Namespace Spotnet

    Friend Class UnCloseableTabItem
        Inherits TabItem

        Shared Sub New()

            DefaultStyleKeyProperty.OverrideMetadata(GetType(UnCloseableTabItem), New FrameworkPropertyMetadata(GetType(UnCloseableTabItem)))

        End Sub

        Public Overrides Sub OnApplyTemplate()

            MyBase.OnApplyTemplate()

        End Sub

    End Class

    Friend Class CloseableTabItem
        Inherits TabItem

        Public AutoSelect As Boolean = True

        Shared Sub New()

            DefaultStyleKeyProperty.OverrideMetadata(GetType(CloseableTabItem), New FrameworkPropertyMetadata(GetType(CloseableTabItem)))

        End Sub

        Public Overrides Sub OnApplyTemplate()

            MyBase.OnApplyTemplate()

            Dim closeButton As Button = TryCast(MyBase.GetTemplateChild("PART_Close"), Button)

            If closeButton IsNot Nothing Then

                closeButton.AddHandler(Button.ClickEvent, New RoutedEventHandler(AddressOf closeButton_Click))

            End If

            If AutoSelect Then
                MyBase.IsSelected = True
            End If

            AutoSelect = False

        End Sub

        Private Sub closeButton_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs)

            CloseMe()

        End Sub

        Public Function CloseMe() As Boolean

            Try

                Dim bErr As Boolean = False
                Dim XX As TabItem = Me
                Dim XXZ As TabControl = Nothing
                Dim ZX As HTMLView = Nothing

                If Not XX Is Nothing Then XXZ = CType(XX.Parent, TabControl)

                If Not XXZ Is Nothing Then

                    If TypeOf XX.Content Is HTMLView Then

                        ZX = CType(XX.Content, HTMLView)

                    End If

                    XXZ.SelectedIndex = 0
                    XXZ.Items.Remove(XX)

                    If My.Settings.SaveTabs Then
                        Dim K3 As DockPanel = CType(XXZ.Parent, DockPanel)
                        Dim K2 As MainWindow = CType(K3.Parent, MainWindow)
                        K2.SaveTabs("")
                    End If

                    If Not ZX Is Nothing Then

                        If Not ZX.Unload() Then
                            bErr = True
                        End If

                    End If

                End If

                Return (Not bErr)

            Catch ex As Exception
                Tools.Foutje("CloseMe: " & ex.Message)
                Return False

            End Try

        End Function

        Private Sub CloseableTabItem_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseDown

            e.Handled = True

            If e.MiddleButton = MouseButtonState.Pressed Or e.RightButton = MouseButtonState.Pressed Then

                If TypeOf e.OriginalSource Is ScrollViewer Then Exit Sub

                closeButton_Click(sender, e)

            End If

        End Sub

    End Class

End Namespace
