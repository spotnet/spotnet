Imports System.Windows.Interop
Imports System.Windows.Media.Imaging
Imports System.Runtime.InteropServices

Public Class Common

    Friend Shared Function GetIcon(ByVal sKey As String) As Image

        Try
            Dim icon As Image = New Image
            icon.Source = New BitmapImage(New Uri("pack://application:,,,/Spotnet;component/Images/" & sKey & ".ico"))
            Return icon
        Catch
            Return Nothing
        End Try

    End Function

    Friend Shared Function GetImage(ByVal sKey As String) As ImageSource

        Try
            Return New BitmapImage(New Uri("pack://application:,,,/Spotnet;component/Images/" & sKey))
        Catch
            Return Nothing
        End Try

    End Function

    Public Shared Sub LaunchBrowser(ByVal sUrl As String)

        Try
            System.Diagnostics.Process.Start(sUrl)
        Catch
            Tools.Foutje("LaunchBrowser:: " & Err.Description)
        End Try

    End Sub

End Class

Public Class Glass

    Structure Margins
        Public Left As Integer
        Public Right As Integer
        Public Top As Integer
        Public Bottom As Integer
    End Structure

    <DllImport("dwmapi.dll")> Public Shared Sub DwmIsCompositionEnabled(ByRef pfEnabled As Boolean)
    End Sub

    Public Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    Public Declare Function DwmExtendFrameIntoClientArea Lib "DwmApi.dll" (ByVal hwnd As IntPtr, ByRef pMarInset As Margins) As Integer

    Dim GlassEnabled As Boolean
    Dim Margin As New Margins

    Dim WM_NCHITTEST As Integer = 132
    Dim HTBORDER As Integer = 18
    Dim HTBOTTOM As Integer = 15
    Dim HTBOTTOMLEFT As Integer = 16
    Dim HTBOTTOMRIGHT As Integer = 17
    Dim HTLEFT As Integer = 10
    Dim HTRIGHT As Integer = 11
    Dim HTTOP As Integer = 12
    Dim HTTOPLEFT As Integer = 13
    Dim HTTOPRIGHT As Integer = 14

    Friend Sub New()

        MyBase.New()

    End Sub

    Private Function WndProc(ByVal hwnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr, ByRef handled As Boolean) As IntPtr

        If (msg = WM_NCHITTEST) Then
            handled = True
            Dim htLocation As Integer = DefWindowProc(hwnd, msg, wParam, lParam).ToInt32
            Select Case (htLocation)
                Case HTBOTTOM, HTBOTTOMLEFT, HTBOTTOMRIGHT, HTLEFT, HTRIGHT, HTTOP, HTTOPLEFT, HTTOPRIGHT
                    htLocation = HTBORDER
            End Select
            Return New IntPtr(htLocation)
        End If

        Return IntPtr.Zero

    End Function

    Public Sub Init(ByRef Form As Window)

        Try

            Dim mainWindowPtr As IntPtr = New WindowInteropHelper(Form).Handle
            Dim mainWindowSrc As HwndSource = HwndSource.FromHwnd(mainWindowPtr)

            mainWindowSrc.AddHook(AddressOf WndProc)

            If Environment.OSVersion.Version.Major >= 6 Then

                DwmIsCompositionEnabled(GlassEnabled)

                If GlassEnabled Then

                    Form.Background = Brushes.Transparent
                    mainWindowSrc.CompositionTarget.BackgroundColor = Colors.Transparent

                    Margin.Left = 1
                    Margin.Right = 1
                    Margin.Bottom = 1
                    Margin.Top = 1

                    Dim hr As Integer = DwmExtendFrameIntoClientArea(mainWindowSrc.Handle, Margin)

                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

End Class