Imports System.Runtime.InteropServices

Class Application

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.

    Private MutX As System.Threading.Mutex
    Private _Splash As SplashScreen = Nothing

    <DllImport("dwmapi.dll")> Public Shared Sub DwmIsCompositionEnabled(ByRef pfEnabled As Boolean)
    End Sub

    Private Sub App_Startup(ByVal sender As Object, ByVal e As StartupEventArgs)

        Try

            Dim bCreated As Boolean = False

            Try
                MutX = New System.Threading.Mutex(False, "Global\" & Spotname, bCreated)
            Catch
            End Try

            If Not bCreated Then
                Foutje("Er loopt al een instantie van Spotnet!", "Fout")
                End
            End If

            _Splash = New SplashScreen("Images/splash.png")
            _Splash.Show(False, True)

            If NeedsTheme() Then
                Dim UR As New Uri("PresentationFramework.Aero;V3.0.0.0;31bf3856ad364e35;component\themes/aero.normalcolor.xaml", UriKind.Relative)
                Resources.MergedDictionaries.Add(CType(Application.LoadComponent(UR), ResourceDictionary))
            End If

        Catch ex As Exception

            Foutje("App_Startup: " & ex.Message)
            End

        End Try

    End Sub

    Private Function NeedsTheme() As Boolean

        If Environment.OSVersion.Version.Major >= 6 Then

            Dim bEnabled As Boolean = False
            DwmIsCompositionEnabled(bEnabled)
            Return Not bEnabled

        Else

            Return True

        End If

    End Function

    Friend Sub CloseSplash(ByVal lDelay As Integer)

        If Not _Splash Is Nothing Then

            Try
                '_Splash.Close(New TimeSpan(0, 0, 0, 0, lDelay))
                _Splash.Close(TimeSpan.Zero)
            Catch
                Try
                    Me.MainWindow.Focus()
                    '_Splash.Close(New TimeSpan(0, 0, 0, 0, lDelay))
                    _Splash.Close(TimeSpan.Zero)
                Catch ex As Exception
                    MsgBox("CloseSplash: " & ex.Message)
                    MsgBox("CloseSplash: " & ex.Message)
                End Try
            End Try

            _Splash = Nothing

        End If

    End Sub

End Class
