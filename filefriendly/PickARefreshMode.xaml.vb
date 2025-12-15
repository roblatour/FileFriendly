Partial Public Class PickARefreshMode

    Private Sub PickARefreshModeWindow_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded

        Me.cbInbox.IsChecked = gRefreshInbox
        Me.cbSent.IsChecked = gRefreshSent
        Me.cbOther.IsChecked = gRefreshAll

    End Sub

    Private Sub Window_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseLeftButtonDown
        DragMove()
    End Sub

    Private Sub BtnOk_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnOK.Click

        gRefreshInbox = Me.cbInbox.IsChecked
        gRefreshSent = Me.cbSent.IsChecked
        gRefreshAll = Me.cbOther.IsChecked
        gRefreshConfirmed = True

        If My.Settings.FirstRun Then
            If gRefreshAll Then

                ShowMessageBox("FileFriendly", _
                  CustomDialog.CustomDialogIcons.Warning, _
                  "As this is the first time you are refreshing 'Other Folders' ...", _
                  "You'll need confirm which ones. You will only be asked to do this once.  Please click [OK] to proceed.", _
                   "You can always change your selection at any time by going to Actions - Options and clicking on the [Scan Folders ...] button.", _
                  "", _
                  CustomDialog.CustomDialogIcons.None, _
                  CustomDialog.CustomDialogButtons.OK, _
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

End Class
