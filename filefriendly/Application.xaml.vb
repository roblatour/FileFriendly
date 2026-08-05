Class Application

    Private _splashScreen As SplashScreen

    Private Sub Application_Startup(ByVal sender As Object, ByVal e As System.Windows.StartupEventArgs) Handles Me.Startup

        Dim mainWindow As New MainWindow()
        Me.MainWindow = mainWindow

        If My.Settings.ShowSplashScreen Then
            mainWindow.AllowsTransparency = True
            mainWindow.Opacity = 0
            mainWindow.Left = -10000
            mainWindow.Top = -10000
            _splashScreen = New SplashScreen()
            PositionSplashScreen(_splashScreen, mainWindow)
            _splashScreen.Show()
        End If

        mainWindow.Show()

    End Sub

    Private Sub PositionSplashScreen(ByVal splashScreen As SplashScreen, ByVal mainWindow As MainWindow)

        Dim mainBounds As System.Windows.Rect = AdjustWindowRect(My.Settings.MainLeft, My.Settings.MainTop, My.Settings.MainWidth, My.Settings.MainHeight, mainWindow.MinWidth, mainWindow.MinHeight)
        Dim folderBounds As System.Windows.Rect = AdjustWindowRect(My.Settings.FoldersLeft, My.Settings.FoldersTop, My.Settings.FoldersWidth, My.Settings.FoldersHeight, 385, 650)

        Dim left As Double = System.Math.Min(mainBounds.Left, folderBounds.Left)
        Dim top As Double = System.Math.Min(mainBounds.Top, folderBounds.Top)
        Dim right As Double = System.Math.Max(mainBounds.Right, folderBounds.Right)
        Dim bottom As Double = System.Math.Max(mainBounds.Bottom, folderBounds.Bottom)

        splashScreen.Left = left + ((right - left - splashScreen.Width) / 2)
        splashScreen.Top = top + ((bottom - top - splashScreen.Height) / 2)

    End Sub

    Public ReadOnly Property IsSplashScreenVisible As Boolean
        Get
            Return _splashScreen IsNot Nothing AndAlso _splashScreen.IsLoaded
        End Get
    End Property

    Public Sub CloseSplashScreen()

        If _splashScreen Is Nothing Then
            Return
        End If

        If _splashScreen.IsLoaded Then
            _splashScreen.Close()
        End If
        _splashScreen = Nothing

        If Me.MainWindow IsNot Nothing Then
            Me.MainWindow.Opacity = 1
            Me.MainWindow.Visibility = Visibility.Visible
        End If

        If gPickAFolderWindow IsNot Nothing Then
            gPickAFolderWindow.Opacity = 1
            If Not gPickAFolderWindow.IsVisible Then
                gPickAFolderWindow.Show()
            End If
        End If

    End Sub

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.

    Private Sub Application_DispatcherUnhandledException(ByVal sender As Object, ByVal e As System.Windows.Threading.DispatcherUnhandledExceptionEventArgs) Handles Me.DispatcherUnhandledException

        ' Quick special-case: ignore layout "Width and Height must be non-negative."
        ' so you can run the app and diagnose in context.
        If e IsNot Nothing AndAlso e.Exception IsNot Nothing AndAlso
       TypeOf e.Exception Is System.ArgumentException AndAlso
       e.Exception.Message IsNot Nothing AndAlso
       e.Exception.Message.IndexOf("Width and Height must be non-negative.", StringComparison.OrdinalIgnoreCase) >= 0 Then

            ' Optionally log it somewhere:
            Console.WriteLine(e.Exception.ToString())

            ' Mark as handled so WPF doesn't crash the app.
            e.Handled = True
            Return
        End If

        Dim outlookVersion As String = ""

        Try
            Dim app As Object = Nothing
            Dim ns As Object = Nothing

            Try
                app = CreateObject("Outlook.Application")
                ns = app.GetNamespace("MAPI")
                outlookVersion = CStr(app.Version)
            Catch
                outlookVersion = ""
            Finally
                ns = Nothing
                app = Nothing
            End Try

        Catch ex As Exception
            ' Console.WriteLine(ex.Message)
        End Try

        outlookVersion = If(outlookVersion, "").Trim()

        If outlookVersion.Length = 0 Then

            If My.Settings.SoundAlert Then Beep()
            MsgBox("FileFriendly has encountered a problem and cannot continue." & vbCrLf & vbCrLf &
               "It appears that Microsoft Outlook is not installed on this computer." & vbCrLf & vbCrLf &
               "FileFreindly requires Outlook to be able to run.",
               MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "FileFriendly - Critical Error")

        Else

            If My.Settings.SoundAlert Then Beep()
            ' Show full exception text instead of only InnerException
            If MsgBox("FileFriendly has encountered a problem and cannot continue." & vbCrLf & vbCrLf &
                  "Would you like to see more detailed information about this problem?",
                  MsgBoxStyle.Critical Or MsgBoxStyle.YesNo, "FileFriendly - Critical Error") = MsgBoxResult.Yes Then

                Dim details As String = If(e.Exception IsNot Nothing,
                                       e.Exception.ToString(),
                                       "")

                MsgBox("Outlook Version: " & outlookVersion & vbCrLf & vbCrLf &
                   "Details:" & vbCrLf & details,
                   MsgBoxStyle.Information, "FileFriendly - Problem Details")
            End If

        End If

        ' Let WPF shut the app down for real fatal errors
        End

    End Sub

End Class
