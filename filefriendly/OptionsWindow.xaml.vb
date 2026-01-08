Partial Public Class OptionsWindow

    Private InitializationComplete As Boolean = False

    Private lDateChoiceAtStartupIsWhenSent As Boolean


    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub



    'can't databind radio boxes, the following is a work around
    Private Sub OptionsWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        Me.rbDockLeft.IsChecked = My.Settings.DockLeft
        Me.rbWhenReceived.IsChecked = My.Settings.WhenReceived
        EnableOptionsFolderButtonsNow()
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

    Public Sub SafelyEnableOptionsFolderButtons()
        Call Dispatcher.BeginInvoke(EnableOptionsFolderButtons)
    End Sub
    Private EnableOptionsFolderButtons As New System.Windows.Forms.MethodInvoker(AddressOf EnableOptionsFolderButtonsNow)
    Private Sub EnableOptionsFolderButtonsNow()

        Me.btnFoldersToScan.IsEnabled = gFolderButtonsOnOptionsWindowEnabled
        Me.btnFoldersToViewInFolderWindow.IsEnabled = gFolderButtonsOnOptionsWindowEnabled

    End Sub

    Private Sub cbScanInbox_Unchecked(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles cbScanInbox.Unchecked, cbScanSent.Unchecked, cbScanAllFolders.Unchecked

        If cbScanInbox.IsChecked OrElse cbScanSent.IsChecked OrElse cbScanAllFolders.IsChecked Then
        Else
            ShowMessageBox("FileFriendly", _
                           CustomDialog.CustomDialogIcons.Warning, _
                           "Note!", _
                           "Scan inbox, sent items and folders shouldn`t all be unchecked at the same time.", _
                           "If you uncheck all three then there will be nothing to review!", _
                           "", _
                           CustomDialog.CustomDialogIcons.None, _
                           CustomDialog.CustomDialogButtons.OK, _
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

End Class
