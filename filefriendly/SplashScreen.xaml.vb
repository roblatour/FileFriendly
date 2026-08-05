Partial Public Class SplashScreen

    Private Sub SplashScreen_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        Dim version As String = System.Windows.Forms.Application.ProductVersion
        While version.EndsWith(".0")
            version = version.Remove(version.Length - 2)
        End While
        tbVersion.Text = "v" & version

    End Sub

    Private Sub SplashScreen_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseLeftButtonDown
        DirectCast(Application.Current, Application).CloseSplashScreen()
    End Sub

End Class
