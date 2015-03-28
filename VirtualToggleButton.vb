Imports System.Windows.Controls.Primitives

Namespace DrWPF.Windows.Controls
    Friend NotInheritable Class VirtualToggleButton
#Region "attached properties"

#Region "IsChecked"

        ''' <summary>
        ''' IsChecked Attached Dependency Property
        ''' </summary>
        Public Shared ReadOnly IsCheckedProperty As DependencyProperty = DependencyProperty.RegisterAttached("IsChecked", GetType(Boolean?), GetType(VirtualToggleButton), New FrameworkPropertyMetadata(CType(False, Boolean?), FrameworkPropertyMetadataOptions.BindsTwoWayByDefault Or FrameworkPropertyMetadataOptions.Journal, New PropertyChangedCallback(AddressOf OnIsCheckedChanged)))

        ''' <summary>
        ''' Gets the IsChecked property.  This dependency property 
        ''' indicates whether the toggle button is checked.
        ''' </summary>
        Private Sub New()
        End Sub
        Public Shared Function GetIsChecked(ByVal d As DependencyObject) As Boolean?
            Return CType(d.GetValue(IsCheckedProperty), Boolean?)
        End Function

        ''' <summary>
        ''' Sets the IsChecked property.  This dependency property 
        ''' indicates whether the toggle button is checked.
        ''' </summary>
        Public Shared Sub SetIsChecked(ByVal d As DependencyObject, ByVal value? As Boolean)
            d.SetValue(IsCheckedProperty, value)
        End Sub

        ''' <summary>
        ''' Handles changes to the IsChecked property.
        ''' </summary>
        Private Shared Sub OnIsCheckedChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim pseudobutton As UIElement = TryCast(d, UIElement)
            If pseudobutton IsNot Nothing Then
                Dim newValue? As Boolean = CType(e.NewValue, Boolean?)
                If newValue = True Then
                    RaiseCheckedEvent(pseudobutton)
                ElseIf newValue = False Then
                    RaiseUncheckedEvent(pseudobutton)
                Else
                    RaiseIndeterminateEvent(pseudobutton)
                End If
            End If
        End Sub

#End Region


#Region "IsThreeState"

        ''' <summary>
        ''' IsThreeState Attached Dependency Property
        ''' </summary>
        Public Shared ReadOnly IsThreeStateProperty As DependencyProperty = DependencyProperty.RegisterAttached("IsThreeState", GetType(Boolean), GetType(VirtualToggleButton), New FrameworkPropertyMetadata(CBool(False)))

        ''' <summary>
        ''' Gets the IsThreeState property.  This dependency property 
        ''' indicates whether the control supports two or three states.  
        ''' IsChecked can be set to null as a third state when IsThreeState is true.
        ''' </summary>
        Public Shared Function GetIsThreeState(ByVal d As DependencyObject) As Boolean
            Return CBool(d.GetValue(IsThreeStateProperty))
        End Function

        ''' <summary>
        ''' Sets the IsThreeState property.  This dependency property 
        ''' indicates whether the control supports two or three states. 
        ''' IsChecked can be set to null as a third state when IsThreeState is true.
        ''' </summary>
        Public Shared Sub SetIsThreeState(ByVal d As DependencyObject, ByVal value As Boolean)
            d.SetValue(IsThreeStateProperty, value)
        End Sub

#End Region

#Region "IsVirtualToggleButton"

        ''' <summary>
        ''' IsVirtualToggleButton Attached Dependency Property
        ''' </summary>
        Public Shared ReadOnly IsVirtualToggleButtonProperty As DependencyProperty = DependencyProperty.RegisterAttached("IsVirtualToggleButton", GetType(Boolean), GetType(VirtualToggleButton), New FrameworkPropertyMetadata(CBool(False), New PropertyChangedCallback(AddressOf OnIsVirtualToggleButtonChanged)))

        ''' <summary>
        ''' Gets the IsVirtualToggleButton property.  This dependency property 
        ''' indicates whether the object to which the property is attached is treated as a VirtualToggleButton.  
        ''' If true, the object will respond to keyboard and mouse input the same way a ToggleButton would.
        ''' </summary>
        Public Shared Function GetIsVirtualToggleButton(ByVal d As DependencyObject) As Boolean
            Return CBool(d.GetValue(IsVirtualToggleButtonProperty))
        End Function

        ''' <summary>
        ''' Sets the IsVirtualToggleButton property.  This dependency property 
        ''' indicates whether the object to which the property is attached is treated as a VirtualToggleButton.  
        ''' If true, the object will respond to keyboard and mouse input the same way a ToggleButton would.
        ''' </summary>
        Public Shared Sub SetIsVirtualToggleButton(ByVal d As DependencyObject, ByVal value As Boolean)
            d.SetValue(IsVirtualToggleButtonProperty, value)
        End Sub

        ''' <summary>
        ''' Handles changes to the IsVirtualToggleButton property.
        ''' </summary>
        Private Shared Sub OnIsVirtualToggleButtonChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim element As IInputElement = TryCast(d, IInputElement)
            If element IsNot Nothing Then
                If CBool(e.NewValue) Then
                    AddHandler element.MouseLeftButtonDown, AddressOf OnMouseLeftButtonDown
                    AddHandler element.KeyDown, AddressOf OnKeyDown
                Else
                    RemoveHandler element.MouseLeftButtonDown, AddressOf OnMouseLeftButtonDown
                    RemoveHandler element.KeyDown, AddressOf OnKeyDown
                End If
            End If
        End Sub

#End Region

#End Region

#Region "routed events"

#Region "Checked"

        ''' <summary>
        ''' A static helper method to raise the Checked event on a target element.
        ''' </summary>
        ''' <param name="target">UIElement or ContentElement on which to raise the event</param>
        Friend Shared Function RaiseCheckedEvent(ByVal target As UIElement) As RoutedEventArgs
            If target Is Nothing Then
                Return Nothing
            End If

            Dim args As New RoutedEventArgs()
            args.RoutedEvent = ToggleButton.CheckedEvent
            target.RaiseEvent(args)
            Return args
        End Function

#End Region

#Region "Unchecked"

        ''' <summary>
        ''' A static helper method to raise the Unchecked event on a target element.
        ''' </summary>
        ''' <param name="target">UIElement or ContentElement on which to raise the event</param>
        Friend Shared Function RaiseUncheckedEvent(ByVal target As UIElement) As RoutedEventArgs
            If target Is Nothing Then
                Return Nothing
            End If

            Dim args As New RoutedEventArgs()
            args.RoutedEvent = ToggleButton.UncheckedEvent
            target.RaiseEvent(args)
            Return args
        End Function

#End Region

#Region "Indeterminate"

        ''' <summary>
        ''' A static helper method to raise the Indeterminate event on a target element.
        ''' </summary>
        ''' <param name="target">UIElement or ContentElement on which to raise the event</param>
        Friend Shared Function RaiseIndeterminateEvent(ByVal target As UIElement) As RoutedEventArgs
            If target Is Nothing Then
                Return Nothing
            End If

            Dim args As New RoutedEventArgs()
            args.RoutedEvent = ToggleButton.IndeterminateEvent
            target.RaiseEvent(args)
            Return args
        End Function

#End Region

#End Region

#Region "private methods"

        Private Shared Sub OnMouseLeftButtonDown(ByVal sender As Object, ByVal e As MouseButtonEventArgs)
            e.Handled = True
            UpdateIsChecked(TryCast(sender, DependencyObject))
        End Sub

        Private Shared Sub OnKeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
            If e.OriginalSource Is sender Then
                If e.Key = Key.Space Then
                    ' ignore alt+space which invokes the system menu
                    If (Keyboard.Modifiers And ModifierKeys.Alt) = ModifierKeys.Alt Then
                        Return
                    End If

                    UpdateIsChecked(TryCast(sender, DependencyObject))
                    e.Handled = True

                ElseIf e.Key = Key.Enter AndAlso CBool((TryCast(sender, DependencyObject)).GetValue(KeyboardNavigation.AcceptsReturnProperty)) Then
                    UpdateIsChecked(TryCast(sender, DependencyObject))
                    e.Handled = True
                End If
            End If
        End Sub

        Private Shared Sub UpdateIsChecked(ByVal d As DependencyObject)
            Dim isChecked? As Boolean = GetIsChecked(d)
            If isChecked = True Then SetIsChecked(d, If(GetIsThreeState(d), True, False)) Else SetIsChecked(d, isChecked.HasValue)
        End Sub

        Private Shared Sub [RaiseEvent](ByVal target As DependencyObject, ByVal args As RoutedEventArgs)
            If TypeOf target Is UIElement Then
                TryCast(target, UIElement).RaiseEvent(args)
            ElseIf TypeOf target Is ContentElement Then
                TryCast(target, ContentElement).RaiseEvent(args)
            End If
        End Sub

#End Region
    End Class
End Namespace