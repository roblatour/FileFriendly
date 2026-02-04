' Copyright Rob Latour, 2026

Partial Public Class PickARefreshMode

    Private Sub PickARefreshModeWindow_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded

        Me.cbInbox.IsChecked = gRefreshInbox
        Me.cbSent.IsChecked = gRefreshSent
        Me.cbOther.IsChecked = gRefreshOtherFolders

    End Sub

    Private Sub Window_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseLeftButtonDown
        DragMove()
    End Sub

    Private Sub BtnOk_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnOK.Click

        gRefreshInbox = Me.cbInbox.IsChecked
        gRefreshSent = Me.cbSent.IsChecked
        gRefreshOtherFolders = Me.cbOther.IsChecked
        gRefreshConfirmed = True

        If My.Settings.FirstRun Then
            If gRefreshOtherFolders Then

                ShowMessageBox("FileFriendly",
                  CustomDialog.CustomDialogIcons.Warning,
                  "As this is the first time you are refreshing 'Other Folders' ...",
                  "You'll need confirm which ones. You will only be asked to do this once.  Please click [OK] to proceed.",
                   "You can always change your selection at any time by going to Actions - Options and clicking on the [Scan Folders ...] button.",
                  "",
                  CustomDialog.CustomDialogIcons.None,
                  CustomDialog.CustomDialogButtons.OK,
                  CustomDialog.CustomDialogResults.OK)

                Me.Hide()

                gFolderReviewWindowContext = FolderReviewContext.ForScanning
                gFolderReviewWindow = New FolderReviewWindow
                gFolderReviewWindow.ShowDialog()
                gFolderReviewWindow = Nothing

            End If
        End If

        Me.Close()

    End Sub

    Private Sub imgClose_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles imgClose.MouseDown
        gRefreshConfirmed = False
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnCancel.Click
        gRefreshConfirmed = False
        Me.Close()
    End Sub

    Private Sub PickARefreshMode_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Me.PreviewKeyDown
        If e.Key <> System.Windows.Input.Key.Left AndAlso e.Key <> System.Windows.Input.Key.Right Then Return

        Dim buttons As New List(Of System.Windows.Controls.Button)

        If btnOK.IsVisible AndAlso btnOK.IsEnabled Then buttons.Add(btnOK)
        If btnCancel.IsVisible AndAlso btnCancel.IsEnabled Then buttons.Add(btnCancel)

        If buttons.Count < 2 Then Return

        Dim focusedButton As System.Windows.Controls.Button = TryCast(System.Windows.Input.Keyboard.FocusedElement, System.Windows.Controls.Button)
        Dim currentIndex As Integer = buttons.IndexOf(focusedButton)

        If currentIndex = -1 Then
            buttons(0).Focus()
            e.Handled = True
            Return
        End If

        Dim nextIndex As Integer
        If e.Key = System.Windows.Input.Key.Left Then
            nextIndex = (currentIndex - 1 + buttons.Count) Mod buttons.Count
        Else
            nextIndex = (currentIndex + 1) Mod buttons.Count
        End If

        buttons(nextIndex).Focus()
        e.Handled = True
    End Sub

End Class
