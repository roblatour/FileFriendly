' Copyright Rob Latour, 2026

Partial Public Class OptionsWindow

    Private InitializationComplete As Boolean = False

    Private lDateChoiceAtStartupIsWhenSent As Boolean
    Private lKeepHiddenEmailsHiddenCurrent As Boolean
    Private EnableOptionsFolderButtons As New System.Windows.Forms.MethodInvoker(AddressOf EnableOptionsFolderButtonsNow)


    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub SafelyEnableOptionsFolderButtons()
        Call Dispatcher.BeginInvoke(EnableOptionsFolderButtons)
    End Sub


    'can't databind radio boxes, the following is a work around
    Private Sub OptionsWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        Me.rbDockLeft.IsChecked = My.Settings.DockLeft
        Me.rbWhenReceived.IsChecked = My.Settings.WhenReceived
        EnableOptionsFolderButtonsNow()

        lKeepHiddenEmailsHiddenCurrent = My.Settings.KeepHiddenEmailsHidden
        InitializationComplete = True

        lDateChoiceAtStartupIsWhenSent = My.Settings.WhenSent

    End Sub

    Private Sub OptionsWindow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing

        If (lDateChoiceAtStartupIsWhenSent = My.Settings.WhenSent) Then
        Else
            gARefreshIsRequired = True
        End If

    End Sub

    Private Sub rbWhenReceived_UnChecked(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles rbWhenReceived.Unchecked
        My.Settings.WhenReceived = False
        My.Settings.WhenSent = True
    End Sub
    Private Sub rbWhenReceived_Checked(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles rbWhenReceived.Checked
        My.Settings.WhenReceived = True
        My.Settings.WhenSent = False
    End Sub

    Private Sub rbDockLeft_Unchecked(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles rbDockLeft.Unchecked
        My.Settings.DockLeft = False
        My.Settings.DockRight = True
        ApplyDocking()
    End Sub
    Private Sub rbDockLeft_Checked(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles rbDockLeft.Checked
        My.Settings.DockLeft = True
        My.Settings.DockRight = False
        ApplyDocking()
    End Sub
    Private Sub ApplyDocking()
        If gWindowDocked Then
            If gPickAFolderWindow IsNot Nothing Then gPickAFolderWindow.SafelyMovePickAFolderWindow()
        End If
    End Sub

    Private Sub cbKeepHiddenEmailsHidden_Changed(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cbKeepHiddenEmailsHidden.Checked, cbKeepHiddenEmailsHidden.Unchecked

        If Not InitializationComplete Then Return

        Dim newValue As Boolean = cbKeepHiddenEmailsHidden.IsChecked

        If newValue = lKeepHiddenEmailsHiddenCurrent Then Return

        If newValue Then

            ' option was selected to persist hidden items

            lKeepHiddenEmailsHiddenCurrent = newValue
            My.Settings.KeepHiddenEmailsHidden = newValue

        Else
            ' option was selected to not persist hidden items

            Dim result = ShowMessageBox("FileFriendly - Confirmation",
                                  CustomDialog.CustomDialogIcons.Question,
                                  "Are you sure?",
                                  "Unchecking this option means that all hidden items will re-appear when FileFriendly is next started.",
                                  "",
                                  "",
                                  CustomDialog.CustomDialogIcons.None,
                                  CustomDialog.CustomDialogButtons.YesNo,
                                  CustomDialog.CustomDialogResults.No)

            If result = CustomDialog.CustomDialogResults.Yes Then

                lKeepHiddenEmailsHiddenCurrent = newValue
                My.Settings.KeepHiddenEmailsHidden = newValue

            Else

                cbKeepHiddenEmailsHidden.IsChecked = lKeepHiddenEmailsHiddenCurrent
                My.Settings.KeepHiddenEmailsHidden = lKeepHiddenEmailsHiddenCurrent

            End If

        End If

    End Sub

    Private Sub Window_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseLeftButtonDown
        DragMove()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnOK.Click
        My.Settings.Save()
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnCancel.Click
        My.Settings.Reload()
        Me.Close()
    End Sub

    Private Sub imgClose_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles imgClose.MouseDown
        My.Settings.Reload()
        Me.Close()
    End Sub

    Private Sub EnableOptionsFolderButtonsNow()

        Me.btnFoldersToScan.IsEnabled = gFolderButtonsOnOptionsWindowEnabled
        Me.btnFoldersToViewInFolderWindow.IsEnabled = gFolderButtonsOnOptionsWindowEnabled

    End Sub

    Private Sub cbScanInbox_Unchecked(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cbScanInbox.Unchecked, cbScanSent.Unchecked, cbScanAllFolders.Unchecked

        If cbScanInbox.IsChecked OrElse cbScanSent.IsChecked OrElse cbScanAllFolders.IsChecked Then
        Else
            ShowMessageBox("FileFriendly",
                           CustomDialog.CustomDialogIcons.Warning,
                           "Note!",
                           "Scan inbox, sent items and folders shouldn`t all be unchecked at the same time.",
                           "If you uncheck all three then there will be nothing to review!",
                           "",
                           CustomDialog.CustomDialogIcons.None,
                           CustomDialog.CustomDialogButtons.OK,
                           CustomDialog.CustomDialogResults.OK)
        End If

    End Sub

    Private Sub btnFoldersToViewInFolderWindow_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnFoldersToViewInFolderWindow.Click

        gFolderReviewWindowContext = FolderReviewContext.ForViewing
        gFolderReviewWindow = New FolderReviewWindow
        gFolderReviewWindow.ShowDialog()
        gFolderReviewWindow = Nothing

        If gPickAFolderWindow IsNot Nothing Then
            gPickAFolderWindow.SafelyRefreshPickAFolderWindow()
        End If

    End Sub

    Private Sub btnFoldersToScan_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnFoldersToScan.Click

        gFolderReviewWindowContext = FolderReviewContext.ForScanning
        gFolderReviewWindow = New FolderReviewWindow
        gFolderReviewWindow.ShowDialog()
        gFolderReviewWindow = Nothing

    End Sub

    Private Sub cbUpgradeNofity_Checked(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles cbUpgradeNofity.Checked

        If InitializationComplete Then
            If cbUpgradeNofity.IsChecked Then
                CheckIfNewVersionIsAvailable()
            End If
        End If

    End Sub

    Private Sub OptionsWindow_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Me.PreviewKeyDown
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
