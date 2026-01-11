Partial Public Class FolderReviewWindow

    ' Central, easily editable list of special folder names to exclude for "Recommended".
    ' Comparison is case-insensitive; values are matched against Branch.Name.
    Private Shared ReadOnly RecommendedExcludedFolderNames As String() = {
        "All Folders",
        "Deleted Items",
        "Drafts",
        "Inbox",
        "Junk",
        "Junk E-mail",
        "Junk Email",
        "News Feed",
        "Outbox",
        "RSS Feeds",
        "Sent Items",
        "Sent Mail",
        "Spam",
        "Sync Issues",
        "Trash",
        "Yammer Root"
    }

    ' Folders whose entire sub-tree should be excluded (unchecked) by Recommended.
    ' Case-insensitive match against Branch.Name.
    Private Shared ReadOnly RecommendedExcludedFolderSubtrees As String() = {
        "Sync Issues",
        "Yammer Root"
    }

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

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

    Private Shared strCollection As System.Collections.Specialized.StringCollection

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

            ' Build look ups for excluded single folders and entire sub-trees.
            Dim excluded As New System.Collections.Generic.HashSet(Of String)(
                RecommendedExcludedFolderNames,
                System.StringComparer.OrdinalIgnoreCase)

            Dim excludedSubtrees As New System.Collections.Generic.HashSet(Of String)(
                RecommendedExcludedFolderSubtrees,
                System.StringComparer.OrdinalIgnoreCase)

            ApplyRecommendedSelection(root, excluded, excludedSubtrees, False)

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

    Private Shared Sub ApplyRecommendedSelection(ByVal branch As TreeViewWithCheckBoxes.FooViewModel,
                                                 ByVal excluded As System.Collections.Generic.HashSet(Of String),
                                                 ByVal excludedSubtrees As System.Collections.Generic.HashSet(Of String),
                                                 ByVal parentInExcludedSubtree As Boolean)

        Dim folderName As String = branch.Name

        ' If any ancestor is an excluded-sub-tree folder, everything below is unchecked.
        Dim thisIsExcludedSubtreeRoot As Boolean = False
        If folderName IsNot Nothing AndAlso excludedSubtrees.Contains(folderName) Then
            thisIsExcludedSubtreeRoot = True
        End If

        Dim inExcludedSubtree As Boolean = parentInExcludedSubtree OrElse thisIsExcludedSubtreeRoot

        If inExcludedSubtree Then

            ' Under "Sync Issues" or "Yammer Root": always unchecked.
            branch.IsChecked = False

        Else

            ' Normal rule: everything checked except folders in RecommendedExcludedFolderNames.
            If folderName IsNot Nothing AndAlso excluded.Contains(folderName) Then
                branch.IsChecked = False
            Else
                branch.IsChecked = True
            End If

        End If

        For Each child As TreeViewWithCheckBoxes.FooViewModel In branch.Children
            ApplyRecommendedSelection(child, excluded, excludedSubtrees, inExcludedSubtree)
        Next

    End Sub

End Class