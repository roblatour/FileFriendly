Class Application

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
            System.Diagnostics.Debug.WriteLine(e.Exception.ToString())

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
            Console.WriteLine(ex.Message)
        End Try

        outlookVersion = If(outlookVersion, "").Trim()

        If outlookVersion.Length = 0 Then

            MsgBox("FileFriendly has encountered a problem and cannot continue." & vbCrLf & vbCrLf &
               "It appears that Microsoft Outlook is not installed on this computer." & vbCrLf & vbCrLf &
               "FileFreindly requires Outlook to be able to run.",
               MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "FileFriendly - Critical Error")

        Else

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
