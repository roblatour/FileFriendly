' Copyright Rob Latour, 2026
Partial Public Class FolderReviewWindow

    ' Central, easily editable list of special folder names to exclude for "Recommended".
    ' Comparison is case-insensitive; values are matched against Branch.Name.
    ' NOTE: Most default folders (Inbox, Sent, Deleted, Drafts, Junk, Outbox, RSS, Sync Issues)
    ' are now detected via Outlook Interop and work across all language versions.
    ' Only folders without OlDefaultFolders equivalents remain here.
    Private Shared ReadOnly RecommendedExcludedFolderNames As String() = {
        "All Folders",
        "News Feed",
        "Yammer Root"
    }

    ' Folders whose entire sub-tree should be excluded (unchecked) by Recommended.
    ' Case-insensitive match against Branch.Name.
    Private Shared ReadOnly RecommendedExcludedFolderSubtrees As String() = {
        "Sync Issues",
        "Yammer Root"
    }

    Private Shared strCollection As System.Collections.Specialized.StringCollection

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Shared Sub DumpTreeView(ByVal Branch As TreeViewWithCheckBoxes.FooViewModel)
        ' uses recursion (I'm so proud)

        If Branch.IsChecked Then
        Else
            strCollection.Add(Branch.FullPathName)
        End If

        For Each Child As TreeViewWithCheckBoxes.FooViewModel In Branch.Children
            DumpTreeView(Child)
        Next

    End Sub

    Private Shared Function IsFolderExcludedByDefault(ByVal fullPathName As String) As Boolean

        If String.IsNullOrEmpty(fullPathName) Then
            Return False
        End If

        Dim idx As Integer = LookupFolderNamesTableIndex(fullPathName)
        If idx < 0 OrElse idx >= gFolderTable.Length Then
            Return False
        End If

        Dim entryId As String = gFolderTable(idx).EntryID
        If String.IsNullOrEmpty(entryId) Then
            Return False
        End If

        If MainWindow.gDefaultInboxEntryIDs.Contains(entryId) Then Return True
        If MainWindow.gDefaultSentEntryIDs.Contains(entryId) Then Return True
        If MainWindow.gDefaultDeletedEntryIDs.Contains(entryId) Then Return True
        If MainWindow.gDefaultDraftsEntryIDs.Contains(entryId) Then Return True
        If MainWindow.gDefaultJunkEntryIDs.Contains(entryId) Then Return True
        If MainWindow.gDefaultOutboxEntryIDs.Contains(entryId) Then Return True
        If MainWindow.gDefaultRssFeedsEntryIDs.Contains(entryId) Then Return True
        If MainWindow.gDefaultSyncIssuesEntryIDs.Contains(entryId) Then Return True

        Return False

    End Function

    Private Shared Sub ApplyRecommendedSelection(ByVal branch As TreeViewWithCheckBoxes.FooViewModel,
                                                 ByVal excluded As System.Collections.Generic.HashSet(Of String),
                                                 ByVal excludedSubtrees As System.Collections.Generic.HashSet(Of String),
                                                 ByVal parentInExcludedSubtree As Boolean)

        Dim folderName As String = branch.Name

        If Not branch.IsEnabled Then
            branch.IsChecked = False
            For Each child As TreeViewWithCheckBoxes.FooViewModel In branch.Children
                ApplyRecommendedSelection(child, excluded, excludedSubtrees, True)
            Next
            Return
        End If

        Dim thisIsExcludedSubtreeRoot As Boolean = False
        If folderName IsNot Nothing AndAlso excludedSubtrees.Contains(folderName) Then
            thisIsExcludedSubtreeRoot = True
        End If

        Dim inExcludedSubtree As Boolean = parentInExcludedSubtree OrElse thisIsExcludedSubtreeRoot

        If inExcludedSubtree Then

            branch.IsChecked = False

        Else

            Dim isExcluded As Boolean = False

            If IsFolderExcludedByDefault(branch.FullPathName) Then
                isExcluded = True

            ElseIf ((gFolderReviewWindowContext = FolderReviewContext.ForScanning) AndAlso (branch.FullPathName.EndsWith("\Archive") OrElse branch.FullPathName.EndsWith("\Archiv"))) Then
                isExcluded = True

            ElseIf folderName IsNot Nothing AndAlso excluded.Contains(folderName) Then
                isExcluded = True
            End If

            branch.IsChecked = Not isExcluded

        End If

        For Each child As TreeViewWithCheckBoxes.FooViewModel In branch.Children
            ApplyRecommendedSelection(child, excluded, excludedSubtrees, inExcludedSubtree)
        Next

    End Sub

    Private Sub FolderReviewWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        Try

            Dim root As TreeViewWithCheckBoxes.FooViewModel = TryCast(Me.Tree.Items(0), TreeViewWithCheckBoxes.FooViewModel)
            Me.Tree.Focus()

        Catch ex As Exception
            If My.Settings.SoundAlert Then Beep()
            ShowMessageBox("FileFriendly",
                         CustomDialog.CustomDialogIcons.Stop,
                         "Unexpected Error!",
                         "FileFriendly has encountered an unexpected error.",
                         ex.ToString,
                         "",
                         CustomDialog.CustomDialogIcons.None,
                         CustomDialog.CustomDialogButtons.OK,
                         CustomDialog.CustomDialogResults.OK)
        End Try

    End Sub

    Private Sub FolderReviewWindow_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If gFolderReviewWindowContext = FolderReviewContext.ForScanning Then
            Me.TabItem1.Header = "Folders to scan"
            Me.Rectangle2.Margin = New System.Windows.Thickness(109, 32, 4, 0)
        Else
            Me.TabItem1.Header = "Folders to show in folder window"
            Me.Rectangle2.Margin = New System.Windows.Thickness(205, 32, 4, 0)
        End If

    End Sub

    Private Sub MainWindow_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseLeftButtonDown
        DragMove()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnOK.Click

        strCollection = New System.Collections.Specialized.StringCollection
        DumpTreeView(Me.Tree.Items(0))

        If gFolderReviewWindowContext = FolderReviewContext.ForViewing Then
            My.Settings.ExcludedViewFolders = strCollection
        Else
            My.Settings.ExcludedScanFolders = strCollection
        End If

        My.Settings.Save()
        strCollection = Nothing

        Me.Close()

    End Sub

    Private Sub imgClose_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles imgClose.MouseDown
        Me.Close()
    End Sub

    Private Sub btnRecommended_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnRecommended.Click

        Try

            If Me.Tree Is Nothing OrElse Me.Tree.Items.Count = 0 Then
                Exit Sub
            End If

            Dim root As TreeViewWithCheckBoxes.FooViewModel =
                TryCast(Me.Tree.Items(0), TreeViewWithCheckBoxes.FooViewModel)

            If root Is Nothing Then
                Exit Sub
            End If

            ' When recommending for viewing recommend all folders be viewable
            ' When recommending for scanning recommended specific folders be excluded (like inbox, sent, draft, etc.)

            If gFolderReviewWindowContext = FolderReviewContext.ForScanning Then

                Dim excluded As New System.Collections.Generic.HashSet(Of String)(
                RecommendedExcludedFolderNames,
                System.StringComparer.OrdinalIgnoreCase)

                Dim excludedSubtrees As New System.Collections.Generic.HashSet(Of String)(
                RecommendedExcludedFolderSubtrees,
                System.StringComparer.OrdinalIgnoreCase)

                ApplyRecommendedSelection(root, excluded, excludedSubtrees, False)

            End If

        Catch ex As Exception
            If My.Settings.SoundAlert Then Beep()
            ShowMessageBox("FileFriendly",
                         CustomDialog.CustomDialogIcons.Stop,
                         "Unexpected Error!",
                         "FileFriendly has encountered an unexpected error.",
                         ex.ToString,
                         "",
                         CustomDialog.CustomDialogIcons.None,
                         CustomDialog.CustomDialogButtons.OK,
                         CustomDialog.CustomDialogResults.OK)
        End Try

    End Sub

    Private Sub FolderReviewWindow_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Me.PreviewKeyDown
        If e.Key <> System.Windows.Input.Key.Left AndAlso e.Key <> System.Windows.Input.Key.Right Then Return

        Dim buttons As New List(Of System.Windows.Controls.Button)

        If btnCancel.IsVisible AndAlso btnCancel.IsEnabled Then buttons.Add(btnCancel)
        If btnRecommended.IsVisible AndAlso btnRecommended.IsEnabled Then buttons.Add(btnRecommended)
        If btnOK.IsVisible AndAlso btnOK.IsEnabled Then buttons.Add(btnOK)

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