Partial Public Class PickAFolder
    Private Const ForwardSlashCharacter As Char = "\"c
    Private OriginalGuidelineHeight As Double

    Private OriginalGrid1Margin As System.Windows.Thickness
    Private OriginalRectangle2Margin As System.Windows.Thickness

    Private ResizeBy As Integer = 0

    Private WindowDockingInProgress As Boolean = False

    Private QuickFilterWord As String = ""
    Private Const BackspaceTag As String = "*Backspace*"
    Private Const AbcTag As String = "*ABC*"
    Private Const SpaceTag As String = "*Space*"
    Private Const BackspaceGlyph As String = "←"
    Private _numericMode As Boolean = False
    Private _lastButton As Button = Nothing

    ' Flag used to suppress TreeView selection handling while rebuilding items
    Private _isLoadingTreeView As Boolean = False

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Try

            OriginalGrid1Margin = Grid1.Margin
            OriginalRectangle2Margin = SeperatorLine.Margin

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

        _lastButton = Button00
        ApplyAlphabetLayout()
        UpdateButton(QuickFilterWord)

    End Sub

    Public Sub SafelyRefreshPickAFolderWindow()
        Call Dispatcher.BeginInvoke(RefreshPickAFolderWindow)
    End Sub
    Private RefreshPickAFolderWindow As New System.Windows.Forms.MethodInvoker(AddressOf RefreshPickAFolderWindowNow)
    Private Sub RefreshPickAFolderWindowNow()
        If QuickFilterWord.Length = 0 Then
            LoadTreeView("None")
        Else
            LoadTreeView(QuickFilterWord)
        End If
    End Sub

    Public Sub SafelyHidePickAFolderWindow()
        Call Dispatcher.BeginInvoke(HidePickAFolderWindow)
    End Sub
    Private HidePickAFolderWindow As New System.Windows.Forms.MethodInvoker(AddressOf HidePickAFolderWindowNow)
    Private Sub HidePickAFolderWindowNow()
        Me.Hide()
    End Sub

    Public Sub SafelyShowPickAFolderWindow()
        Call Dispatcher.BeginInvoke(ShowPickAFolderWindow)
    End Sub
    Private ShowPickAFolderWindow As New System.Windows.Forms.MethodInvoker(AddressOf ShowPickAFolderWindowNow)
    Private Sub ShowPickAFolderWindowNow()
        Me.Show()
    End Sub


    Public Sub SafelyMakePickAFolderWindowTopMost()
        Call Dispatcher.BeginInvoke(MakePickAFolderWindowTopMost)
    End Sub
    Private MakePickAFolderWindowTopMost As New System.Windows.Forms.MethodInvoker(AddressOf MakePickAFolderWindowTopMostNow)

    Private Sub MakePickAFolderWindowTopMostNow()

        Try

            gPickAFolderWindow.BringIntoView()
            gPickAFolderWindow.Focus()

            If gMainWindowIsMaximized Then

                If gWindowDocked Then
                    DockUndockWindow("UnDock")
                    WindowDockingInProgress = False
                End If
                MakeTopMost(True, PickAWindowHandle)

            Else

                If gPickAFolderWindowWasDocedWhenMainWindowWasMaximimized Then
                    DockUndockWindow("Dock")
                    WindowDockingInProgress = False
                End If
                MakeTopMost(False, PickAWindowHandle)

            End If

            If Me.WindowState = Windows.WindowState.Minimized Then
                Me.WindowState = Windows.WindowState.Normal
            End If

        Catch ex As Exception

        End Try

    End Sub

    Public Sub SafelyMovePickAFolderWindow()
        Call Dispatcher.BeginInvoke(MovePickAFolderWindow)
    End Sub
    Private MovePickAFolderWindow As New System.Windows.Forms.MethodInvoker(AddressOf MovePickAFolderWindowNow)
    Private Sub MovePickAFolderWindowNow()

        PlaceWindow()

    End Sub

    Private Sub PickAFolder_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If gMinimizedAtEarlyStartup Then
            Me.Hide()
            gMinimizedAtEarlyStartup = False
        End If

        'ok so here's the deal
        ' clicking on this form or on the main window form needs to bring both to the forefront
        ' I do this by first bringing the other from to the forefront and then this one
        ' trick is the other form does the same
        ' problem is we get into a bit of a loop when each form trigger the activation of another
        ' the following flip flop code ensure that rather then a loop, each form is only brought forward once

        Static Dim FlipFlop As Boolean = True
        If FlipFlop Then
            gMainWindow.Activate()
            gMainWindow.BringIntoView()
            FlipFlop = False
        Else
            FlipFlop = True
        End If

    End Sub

    Private Sub Window_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles Me.MouseEnter
        Try
            ' If the main window is active and mouse enters Pick A Folder, activate Pick A Folder
            Dim main As MainWindow = TryCast(System.Windows.Application.Current.MainWindow, MainWindow)
            If main IsNot Nothing AndAlso main.IsActive AndAlso Not Me.IsActive Then
                Me.Activate()
            End If
        Catch ex As Exception
            ' Swallow or log as per your existing style
        End Try
    End Sub

    Private Sub PickAFolder_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        Try

            Dim foldersPlacement As System.Windows.Rect = AdjustWindowRect(My.Settings.FoldersLeft, My.Settings.FoldersTop, My.Settings.FoldersWidth, My.Settings.FoldersHeight, Me.MinWidth, Me.MinHeight)

            Me.Width = foldersPlacement.Width

            OriginalGuidelineHeight = Me.Guideline.ActualHeight

            ResetLookOfWindow()
            ApplyAlphabetLayout()

            If gWindowDocked Then
                Dim dockHeight As Double = If(gmwHeight > 0, gmwHeight, foldersPlacement.Height)
                Me.Height = dockHeight
                PlaceWindow()
                Dim adjustedDocked As System.Windows.Rect = AdjustWindowRect(Me.Left, Me.Top, Me.ActualWidth, Me.ActualHeight, Me.MinWidth, Me.MinHeight)
                Me.Width = adjustedDocked.Width
                Me.Height = adjustedDocked.Height
                Me.Left = adjustedDocked.X
                Me.Top = adjustedDocked.Y
            Else
                Me.Height = foldersPlacement.Height
                Me.Left = foldersPlacement.X
                Me.Top = foldersPlacement.Y
            End If

            If My.Settings.SoundDocking Then
                My.Computer.Audio.Play(gDockSound, AudioPlayMode.Background)
            End If

            Dim hwndSource As System.Windows.Interop.HwndSource = TryCast(PresentationSource.FromVisual(Me), System.Windows.Interop.HwndSource)
            If hwndSource IsNot Nothing Then
                PickAWindowHandle = hwndSource.Handle
            End If

            DockUndockWindow("Initial Load")

        Catch ex As Exception

        End Try


    End Sub

    Public intRecommendation1 As Integer = -1
    Public intRecommendation2 As Integer = -1
    Public intRecommendation3 As Integer = -1
    Public intRecommendation4 As Integer = -1

    Public Sub UpdateRecommendationsOnPickAFolderWindow()
        Call Dispatcher.BeginInvoke(UpdateRecommendations)
    End Sub
    Private UpdateRecommendations As New System.Windows.Forms.MethodInvoker(AddressOf ResetLookOfWindow)

    Private Sub ResetLookOfWindow()

        PopulateButtonContents(-2) 'reset

        PopulateButtonContents(intRecommendation3)
        PopulateButtonContents(intRecommendation2)
        PopulateButtonContents(intRecommendation1)

        If TreeView1.SelectedItem IsNot Nothing Then
            Dim tvi As TreeViewItem = TryCast(TreeView1.SelectedItem, TreeViewItem)
            If tvi IsNot Nothing AndAlso tvi.Tag IsNot Nothing Then
                Dim tagText As String = TryCast(tvi.Tag, String)
                If tagText IsNot Nothing Then
                    intRecommendation4 = LookupFolderNamesTableIndex(tagText)
                End If
            End If
        End If
        PopulateButtonContents(intRecommendation4)

        'Dynamically Resize window contents depending on how many recommendations there are

        Grid1.Margin = OriginalGrid1Margin
        SeperatorLine.Margin = OriginalRectangle2Margin

        ResizeBy = 0
        If btnPick4.Visibility = Windows.Visibility.Hidden Then ResizeBy += CInt(btnPick1.Height) + 1
        If btnPick3.Visibility = Windows.Visibility.Hidden Then ResizeBy += CInt(btnPick1.Height) + 1
        If btnPick2.Visibility = Windows.Visibility.Hidden Then ResizeBy += CInt(btnPick1.Height) + 1
        If btnPick1.Visibility = Windows.Visibility.Hidden Then ResizeBy += CInt(btnPick1.Height) + 1
        If ResizeBy > 0 Then
            Me.Grid1.Margin = New System.Windows.Thickness(Me.Grid1.Margin.Left, Me.Grid1.Margin.Top, Me.Grid1.Margin.Right, Me.Grid1.Margin.Bottom - ResizeBy)
            Me.SeperatorLine.Margin = New System.Windows.Thickness(Me.SeperatorLine.Margin.Left, Me.SeperatorLine.Margin.Top, Me.SeperatorLine.Margin.Right, Me.SeperatorLine.Margin.Bottom - ResizeBy)
        End If

        gMainWindow.SafelyUpdateContextMenu()

    End Sub

    Private Sub PopulateButtonContents(ByVal int As Integer)

        ' -2   reset
        ' -1   no recommendation
        ' >=0  recommended folder

        If int = -1 Then Exit Sub

        If int = -2 Then

            btnPick1.Content = ""
            btnPick2.Content = ""
            btnPick3.Content = ""
            btnPick4.Content = ""
            btnPick1.Visibility = Windows.Visibility.Hidden
            btnPick2.Visibility = Windows.Visibility.Hidden
            btnPick3.Visibility = Windows.Visibility.Hidden
            btnPick4.Visibility = Windows.Visibility.Hidden
            SeperatorLine.Visibility = Windows.Visibility.Hidden

            gContextFile1 = ""
            gContextFile2 = ""
            gContextFile3 = ""
            gContextFile4 = ""

        Else

            Dim str As String = gFolderNamesTable(int).TrimStart(ForwardSlashCharacter)

            If String.IsNullOrEmpty(TryCast(Me.btnPick4.Content, String)) Then

                gContextFile4 = str
                Me.btnPick4.Content = str
                Me.btnPick4.Visibility = Windows.Visibility.Visible
                SeperatorLine.Visibility = Windows.Visibility.Visible


            ElseIf String.IsNullOrEmpty(TryCast(Me.btnPick3.Content, String)) Then

                'ensure no duplicates
                If Me.btnPick4.Content.ToString <> str Then
                    gContextFile3 = str
                    Me.btnPick3.Content = str
                    Me.btnPick3.Visibility = Windows.Visibility.Visible
                End If


            ElseIf String.IsNullOrEmpty(TryCast(Me.btnPick2.Content, String)) Then

                'ensure no duplicates
                If (Me.btnPick4.Content.ToString <> str) And (Me.btnPick3.Content.ToString <> str) Then
                    gContextFile2 = str
                    Me.btnPick2.Content = str
                    Me.btnPick2.Visibility = Windows.Visibility.Visible
                End If


            ElseIf String.IsNullOrEmpty(TryCast(Me.btnPick1.Content, String)) Then
                'ensure no duplicates
                If (Me.btnPick4.Content.ToString <> str) AndAlso (Me.btnPick3.Content.ToString <> str) AndAlso (Me.btnPick2.Content.ToString <> str) Then
                    gContextFile1 = str
                    Me.btnPick1.Content = str
                    Me.btnPick1.Visibility = Windows.Visibility.Visible
                End If

            End If

        End If

    End Sub

    Private Sub Window_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseLeftButtonDown

        Try
            gWhoIsInControl = WhoIsInControlType.PickAFolder
            DragMove()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub PlaceWindow()

        gWhoIsInControl = WhoIsInControlType.Main
        Me.Height = gmwHeight
        Me.Top = gmwTop
        If My.Settings.DockLeft Then
            Me.Left = gmwLeft - Me.ActualWidth
        Else
            Me.Left = gmwLeft + gmwWidth
        End If
        gWhoIsInControl = WhoIsInControlType.PickAFolder

    End Sub

    Private Sub PushPin_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles PushPin.MouseDown

        If gMainWindowIsMaximized Then
            DockUndockWindow("UnDock")
        Else
            If gWindowDocked Then
                DockUndockWindow("UnDock")
            Else
                DockUndockWindow("Dock")
            End If
        End If

    End Sub

    Private Sub DockUndockWindow(ByVal Action As String)

        Dim imageUri As Uri = Nothing
        Dim BitmapSource As BitmapSource

        Try

            If WindowDockingInProgress Then

                WindowDockingInProgress = False

            Else
                Select Case Action

                    Case Is = "Dock"

                        If gWindowDocked Then Exit Select

                        gWindowDocked = True
                        WindowDockingInProgress = True
                        imageUri = New Uri("/filefriendly;component/Resources/pushpindown.gif", UriKind.Relative)
                        PlaceWindow()

                        Try
                            If My.Settings.SoundDocking Then
                                My.Computer.Audio.Play(gDockSound, AudioPlayMode.Background)
                            End If
                        Catch ex As Exception
                        End Try

                    Case Is = "UnDock"

                        If Not gWindowDocked Then Exit Select

                        gWindowDocked = False
                        imageUri = New Uri("/filefriendly;component/Resources/pushpinup.gif", UriKind.Relative)

                        'nudge the two windows apart

                        Dim ScreenDimensions As System.Drawing.Rectangle = System.Windows.Forms.Screen.GetWorkingArea(ScreenDimensions)
                        Dim WorkAreaHeight As Integer = ScreenDimensions.Height

                        If Me.Top > WorkAreaHeight - 40 Then
                            Me.Top -= 4
                        Else
                            Me.Top += 4
                        End If

                        If Me.Left > 5 Then
                            Me.Left -= 4
                        Else
                            Me.Left += 4

                        End If

                    Case Is = "Initial Load"
                        If gWindowDocked Then
                            imageUri = New Uri("/filefriendly;component/Resources/pushpindown.gif", UriKind.Relative)
                        Else
                            imageUri = New Uri("/filefriendly;component/Resources/pushpinup.gif", UriKind.Relative)
                        End If

                End Select

            End If

            If imageUri IsNot Nothing Then
                BitmapSource = New BitmapImage(imageUri)
                PushPin.Source = BitmapSource
            End If

        Catch ex As Exception

        End Try

        imageUri = Nothing
        BitmapSource = Nothing

    End Sub

    Private Sub btnPick1_Click_1(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnPick1.Click, btnPick2.Click, btnPick3.Click, btnPick4.Click

        Try

            Dim btn As Button = TryCast(sender, Button)
            If btn Is Nothing Then Exit Sub

            Dim contentText As String = TryCast(btn.Content, String)
            If String.IsNullOrEmpty(contentText) Then Exit Sub

            Dim WinningFolderNumber As Integer = LookupFolderNamesTableIndex(contentText)

            If WinningFolderNumber >= 0 Then
                gPickFromContextMenuOverride = WinningFolderNumber
                gMainWindow.SafelyUpdateRecommendationFromPickAFolderWindow()
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

    Private Sub TreeView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles TreeView1.MouseDoubleClick

        If TreeView1.SelectedItem IsNot Nothing Then

            Dim tvi As TreeViewItem = TryCast(TreeView1.SelectedItem, TreeViewItem)
            If tvi Is Nothing OrElse tvi.Tag Is Nothing Then
                Exit Sub
            End If

            Me.btnPick4.Content = tvi.Tag
            intRecommendation4 = LookupFolderNamesTableIndex(Me.btnPick4.Content.ToString)
            gMainWindow.intRecommendationFinal = intRecommendation4.ToString

            tvi.IsExpanded = True

            ResetLookOfWindow()

            gPickFromContextMenuOverride = intRecommendation4
            gMainWindow.SafelyUpdateRecommendationFromPickAFolderWindow()

        End If

    End Sub

    Private Sub TreeView1_SelectedItemChanged(ByVal sender As Object, ByVal e As System.Windows.RoutedPropertyChangedEventArgs(Of Object)) Handles TreeView1.SelectedItemChanged

        ' Ignore selection changes while we are rebuilding the TreeView
        If _isLoadingTreeView Then
            Exit Sub
        End If

        Dim tvi As TreeViewItem = TryCast(TreeView1.SelectedItem, TreeViewItem)
        If tvi Is Nothing OrElse tvi.Tag Is Nothing Then
            Exit Sub
        End If

        Me.btnPick4.Content = tvi.Tag
        intRecommendation4 = LookupFolderNamesTableIndex(Me.btnPick4.Content.ToString)

        ResetLookOfWindow()

    End Sub

    Private Sub MainWindow_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Me.KeyUp
        ProcessKeyUp(e)
    End Sub
    Private Sub MainWindow_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Me.KeyDown
        ProcessKeyDown(e)
    End Sub

    Public Sub SafelyUpdateQuickFilter()
        Call Dispatcher.BeginInvoke(UpdateQuickFilter)
    End Sub
    Private UpdateQuickFilter As New System.Windows.Forms.MethodInvoker(AddressOf UpdateQuickFilterNow)
    Private Sub UpdateQuickFilterNow()

        ProcessIncomingText(gSentText)

    End Sub

    Private Function QuickFilterButtons() As List(Of Button)
        Return New List(Of Button) From {Button00, Button01, Button02, Button03, Button04, Button05, Button06, Button07, Button08, Button09, Button10, Button11, Button12, Button13, Button14, Button15, Button16, Button17, Button18, Button19, Button20, Button21, Button22, Button23, Button24, Button25, Button26, Button27, Button28, Button30, Button29}
    End Function

    Private Sub ResetButtonAppearance(ByVal button As Button)
        If button Is Nothing Then Exit Sub
        Dim tagValue As String = TryCast(button.Tag, String)
        If String.Equals(tagValue, "*None*", StringComparison.Ordinal) Then
            button.Content = "None"
        ElseIf String.Equals(tagValue, BackspaceTag, StringComparison.Ordinal) Then
            button.Content = BackspaceGlyph
        ElseIf String.Equals(tagValue, AbcTag, StringComparison.Ordinal) Then
            button.Content = "ABC..."
        ElseIf String.Equals(tagValue, SpaceTag, StringComparison.Ordinal) Then
            button.Content = "{space}"
        Else
            button.Content = tagValue
        End If
        button.IsEnabled = True
    End Sub

    Private Sub ResetButtonSelection()
        If _lastButton IsNot Nothing Then
            ResetButtonAppearance(_lastButton)
            _lastButton = Nothing
        End If
    End Sub

    Private Function FindButtonByTag(ByVal tagValue As String) As Button
        For Each button In QuickFilterButtons()
            If button IsNot Nothing AndAlso button.Visibility = Windows.Visibility.Visible Then
                Dim storedTag As String = TryCast(button.Tag, String)
                If storedTag IsNot Nothing AndAlso String.Equals(storedTag, tagValue, StringComparison.OrdinalIgnoreCase) Then
                    Return button
                End If
            End If
        Next
        Return Nothing
    End Function

    Private Sub ApplyAlphabetLayout()
        _numericMode = False
        Dim letters As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Button00.IsEnabled = True
        Dim buttons = QuickFilterButtons()
        For i As Integer = 0 To letters.Length - 1
            Dim btn As Button = buttons(i + 1)
            btn.Visibility = Windows.Visibility.Visible
            btn.Content = letters(i)
            btn.Tag = letters(i).ToString
            btn.IsEnabled = True
        Next
        Button27.Visibility = Windows.Visibility.Visible
        Button27.Content = "!"
        Button27.Tag = "!"
        Button27.IsEnabled = True
        Button28.Visibility = Windows.Visibility.Visible
        Button28.Content = "123..."
        Button28.Tag = "9"
        Button28.IsEnabled = True
        Button30.Visibility = Windows.Visibility.Visible
        Button30.Content = "{space}"
        Button30.Tag = SpaceTag
        Button30.IsEnabled = True
        Button29.Visibility = Windows.Visibility.Visible
        Button29.Content = BackspaceGlyph
        Button29.Tag = BackspaceTag
        Button29.IsEnabled = True
        ResetButtonSelection()
    End Sub

    Private Sub EnterNumericMode()
        _numericMode = True
        Button00.IsEnabled = True
        Dim digits As String() = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
        Dim buttons = QuickFilterButtons()
        For i As Integer = 0 To 10
            Dim btn As Button = buttons(i + 1)
            Dim digitValue As String = digits(i Mod digits.Length)
            btn.Visibility = Windows.Visibility.Visible
            btn.Content = digitValue
            btn.Tag = digitValue
            btn.IsEnabled = True
        Next
        Button12.Visibility = Windows.Visibility.Visible
        Button12.Content = "ABC..."
        Button12.Tag = AbcTag
        Button12.IsEnabled = True
        Button13.Visibility = Windows.Visibility.Visible
        Button13.Content = BackspaceGlyph
        Button13.Tag = BackspaceTag
        Button13.IsEnabled = True
        For i As Integer = 14 To buttons.Count - 1
            Dim btn As Button = buttons(i)
            btn.Visibility = Windows.Visibility.Collapsed
        Next
        ResetButtonSelection()
    End Sub

    Private Sub ProcessIncomingText(ByVal IncomingText As String)

        If gScanningFolders Then Exit Sub

        If IncomingText = "Escape" Then
            QuickFilterWord = ""
            UpdateButton(QuickFilterWord)
            LoadTreeView("None")
            Exit Sub
        End If

        If IncomingText = vbBack Then
            If QuickFilterWord.Length > 0 Then
                QuickFilterWord = Microsoft.VisualBasic.Left(QuickFilterWord, QuickFilterWord.Length - 1)
            End If
        Else

            If QuickFilterWord.Length = 0 Then
                Dim ch As Char = IncomingText(0)
                If (ch < "A"c OrElse ch > "Z"c) AndAlso (ch < "a"c OrElse ch > "z"c) AndAlso (ch < "0"c OrElse ch > "9"c) AndAlso (ch <> "!"c) AndAlso (ch <> " "c) Then
                    Exit Sub
                End If
            End If

            QuickFilterWord &= IncomingText
            If QuickFilterWord.Length > 0 Then
                Dim firstChar As Char = QuickFilterWord(0)
                If Char.IsLetter(firstChar) Then
                    QuickFilterWord = Char.ToUpper(firstChar) & QuickFilterWord.Substring(1)
                End If
            End If
        End If

        UpdateButton(QuickFilterWord)

        If QuickFilterWord.Length = 0 Then
            LoadTreeView("None")
        Else
            LoadTreeView(QuickFilterWord)
        End If

    End Sub


    Private Sub Button00_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles _
                Button00.Click, Button01.Click, Button02.Click, Button03.Click, Button04.Click, Button05.Click, Button06.Click, Button07.Click, Button08.Click, Button09.Click,
                Button10.Click, Button11.Click, Button12.Click, Button13.Click, Button14.Click, Button15.Click, Button16.Click, Button17.Click, Button18.Click, Button19.Click,
                Button20.Click, Button21.Click, Button22.Click, Button23.Click, Button24.Click, Button25.Click, Button26.Click, Button27.Click, Button28.Click, Button30.Click, Button29.Click

        If gScanningFolders AndAlso Not Object.ReferenceEquals(sender, Button00) Then Exit Sub

        Dim clickedButton As Button = TryCast(sender, Button)
        If clickedButton Is Nothing Then Exit Sub

        Try

            If clickedButton Is Button00 Then
                QuickFilterWord = ""
                _numericMode = False
                ApplyAlphabetLayout()
                UpdateButton(QuickFilterWord)
                LoadTreeView("None")
                Return
            End If

            If clickedButton Is Button28 Then
                EnterNumericMode()
                UpdateButton(QuickFilterWord)
                Return
            End If

            Dim clickedTag As String = TryCast(clickedButton.Tag, String)

            If String.Equals(clickedTag, AbcTag, StringComparison.Ordinal) Then
                ApplyAlphabetLayout()
                UpdateButton(QuickFilterWord)
                Return
            End If

            If String.Equals(clickedTag, BackspaceTag, StringComparison.Ordinal) Then
                ProcessIncomingText(vbBack)
                Return
            End If

            If String.Equals(clickedTag, SpaceTag, StringComparison.Ordinal) Then
                ProcessIncomingText(" ")
                Return
            End If

            If String.IsNullOrEmpty(clickedTag) Then Return

            ProcessIncomingText(clickedTag)

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


    Private Sub UpdateQuickFilterStatus()

        If QuickFilterStatus IsNot Nothing Then

            ' change the display of the quick filter word below to display it in proper case
            Dim QuickFilterWordProperCase As String = ""

            If Not String.IsNullOrEmpty(QuickFilterWord) Then

                Dim ti As System.Globalization.TextInfo = System.Globalization.CultureInfo.CurrentCulture.TextInfo
                QuickFilterWordProperCase = ti.ToTitleCase(QuickFilterWord.ToLowerInvariant())

            End If

            QuickFilterStatus.Text = "Quick filter: " & QuickFilterWordProperCase
        End If

    End Sub

    Private Sub UpdateButton(ByVal QuickFilterWord As String)

        ResetButtonSelection()
        Button00.IsEnabled = True
        UpdateQuickFilterStatus()

    End Sub

    Private Sub LoadTreeView(Optional ByVal Filter As String = "None")

        If gScanningFolders Then Exit Sub

        ' Suppress selection-changed logic while we clear and repopulate the TreeView
        _isLoadingTreeView = True

        Try

            Dim strCollection = New System.Collections.Specialized.StringCollection
            strCollection = My.Settings.ExcludedViewFolders

            Dim MatchTargetA As String = ForwardSlashCharacter & QuickFilterWord.ToUpper
            Dim MatchTargetB As String = " " & QuickFilterWord.ToUpper
            Dim MatchTargetC As String = "-" & QuickFilterWord.ToUpper
            Dim MatchTargetD As String = QuickFilterWord.ToUpper
            Dim ws As String

            Dim AllEntriesDisplayed As Boolean
            Dim OnlyOneFilteredSection As String = ""

            Try
                TreeView1.Items.Clear()
            Catch ex As Exception
            End Try

            Dim WorkingFoldersNameTable(gFolderTable.Length - 1) As String

            Dim AllFoldersFilteredOut As Boolean = False
            Dim AllFoldersWhereExcludedInOptions As Boolean = False

            'remove exclude folders

            Dim NewWorkingFoldersNameTableLength As Integer

            If strCollection Is Nothing Then

                Array.Copy(gFolderNamesTable, WorkingFoldersNameTable, gFolderNamesTable.Length)
                NewWorkingFoldersNameTableLength = WorkingFoldersNameTable.Length - 1

            Else

                If WorkingFoldersNameTable.Length > 0 Then

                    Dim ii As Integer = 0
                    For i As Integer = 0 To gFolderNamesTable.Length - 1

                        'only copy entries if they are NOT found in the exclusion table
                        If strCollection.IndexOf(gFolderNamesTable(i)) < 0 Then
                            WorkingFoldersNameTable(ii) = gFolderNamesTable(i)
                            ii += 1
                        End If

                    Next

                    NewWorkingFoldersNameTableLength = ii

                    If NewWorkingFoldersNameTableLength = 0 Then
                        AllFoldersWhereExcludedInOptions = True
                    Else
                        NewWorkingFoldersNameTableLength -= 1
                    End If

                End If

            End If

            ReDim Preserve WorkingFoldersNameTable(NewWorkingFoldersNameTableLength)

            If Filter = "None" Then

                QuickFilterWord = ""
                AllEntriesDisplayed = True

                ApplyAlphabetLayout()
                UpdateButton(QuickFilterWord)

            Else

                ' copy only the entries over that match the search criteria

                Dim ii As Integer = 0

                For i = 0 To WorkingFoldersNameTable.Length - 1
                    ws = WorkingFoldersNameTable(i).ToUpper
                    If ws.Contains(MatchTargetA) OrElse ws.Contains(MatchTargetB) OrElse ws.Contains(MatchTargetC) OrElse ws.Contains(MatchTargetD) Then
                        WorkingFoldersNameTable(ii) = WorkingFoldersNameTable(i)
                        ii += 1
                    End If
                Next

                If ii = 0 Then AllFoldersFilteredOut = True
                If ii = 1 Then OnlyOneFilteredSection = WorkingFoldersNameTable(0)

                ReDim Preserve WorkingFoldersNameTable(ii - 1)

                AllEntriesDisplayed = (WorkingFoldersNameTable.Length = gFolderTable.Length)

            End If

            If AllFoldersWhereExcludedInOptions Then

                TreeView1.Items.Add("All folders were excluded in the options")

            ElseIf AllFoldersFilteredOut Then

                TreeView1.Items.Add("There are no folders starting with a word in them starting with '" & QuickFilterWord & "'.")

            Else

                Array.Sort(WorkingFoldersNameTable)

                'make sure there is one, and only one, slash at the end of every entry
                For i = 0 To WorkingFoldersNameTable.Length - 1
                    WorkingFoldersNameTable(i) = WorkingFoldersNameTable(i).TrimEnd("\"c) & ForwardSlashCharacter
                Next

                'Make sure all entries have all portions of their partial paths defined

                Dim FullPath() As String

                Dim OriginalTableLength As Integer = WorkingFoldersNameTable.Length

                For i = 0 To WorkingFoldersNameTable.Length - 1

                    FullPath = WorkingFoldersNameTable(i).Split("\"c)
                    ws = ""

                    For ii As Integer = 0 To FullPath.Length - 2

                        ws = ws & FullPath(ii).Trim & ForwardSlashCharacter

                        If Array.IndexOf(WorkingFoldersNameTable, ws) = -1 Then
                            ReDim Preserve WorkingFoldersNameTable(WorkingFoldersNameTable.Length)
                            WorkingFoldersNameTable(WorkingFoldersNameTable.Length - 1) = ws
                        End If

                    Next

                Next

                Dim NewTableLength As Integer = WorkingFoldersNameTable.Length

                Array.Sort(WorkingFoldersNameTable)

                'Get rid of the 1st two entries ("\" and "\\" ), also get rid of trailing "\"
                For i = 2 To WorkingFoldersNameTable.Length - 1
                    WorkingFoldersNameTable(i - 2) = WorkingFoldersNameTable(i).Trim().TrimEnd("\"c).Trim()
                Next

                ReDim Preserve WorkingFoldersNameTable(WorkingFoldersNameTable.Length - 3)

                'Load the Treeveiw
                Dim Nodes(WorkingFoldersNameTable.Length) As TreeViewItem
                Dim LastNodeAtThisLevel(255) As Integer ' up to a 256 levels of branches

                Dim CurrentFolder() As String
                Dim CurrentFolderName As String
                Dim CurrentLevel As Integer

                For x As Integer = 0 To WorkingFoldersNameTable.Length - 1

                    CurrentFolder = WorkingFoldersNameTable(x).Split("\"c)
                    CurrentLevel = CurrentFolder.Length - 1
                    CurrentFolderName = CurrentFolder(CurrentLevel)
                    CurrentLevel -= 2 ' first two levels are null and need to be ignored

                    Nodes(x) = New TreeViewItem
                    Nodes(x).Header = CurrentFolderName
                    Nodes(x).Tag = WorkingFoldersNameTable(x)
                    Nodes(x).IsExpanded = True

                    If AllEntriesDisplayed Then
                        Nodes(x).Foreground = System.Windows.Media.Brushes.Brown
                    Else
                        ws = CurrentFolderName.ToUpper
                        If (ws.StartsWith(MatchTargetD)) OrElse (ws.Contains(MatchTargetB)) OrElse (ws.Contains(MatchTargetC)) Then
                            Nodes(x).Foreground = System.Windows.Media.Brushes.Black
                        Else
                            Nodes(x).Foreground = System.Windows.Media.Brushes.Brown
                        End If
                    End If

                    LastNodeAtThisLevel(CurrentLevel) = x

                    If CurrentLevel = 0 Then
                        TreeView1.Items.Add(Nodes(x))
                    Else
                        Nodes(LastNodeAtThisLevel(CurrentLevel - 1)).Items.Add(Nodes(x))
                    End If

                Next

                CurrentFolder = Nothing
                LastNodeAtThisLevel = Nothing
                Nodes = Nothing

            End If

            TreeView1.IsEnabled = True
            TreeView1.IsHitTestVisible = True

            If OnlyOneFilteredSection.Length > 0 Then
                Me.btnPick4.Content = OnlyOneFilteredSection
                intRecommendation4 = LookupFolderNamesTableIndex(OnlyOneFilteredSection)
            End If

            intRecommendation1 = -1
            intRecommendation2 = -1
            intRecommendation3 = -1
            intRecommendation4 = -1

            ResetLookOfWindow()

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
        Finally
            _isLoadingTreeView = False
        End Try

    End Sub

    Private Sub PickAFolder_SizeChanged(ByVal sender As Object, ByVal e As System.Windows.SizeChangedEventArgs) Handles Me.SizeChanged

        ' there are 30 objects within grid2 - 1 label and 29 buttons

        ' do not resize the label
        ' resize the 29 buttons
        ' a guideline is used to set the target height of the buttons

        Dim NewHeight As Double = (Guideline.ActualHeight / QuickFilterButtons().Count)

        For Each obj As Object In Grid2.Children
            Dim btn As Button = TryCast(obj, Button)
            If btn IsNot Nothing Then
                btn.Height = NewHeight
            End If
        Next

        Me.btnPick1.Width = SeperatorLine.ActualWidth
        Me.btnPick2.Width = SeperatorLine.ActualWidth
        Me.btnPick3.Width = SeperatorLine.ActualWidth
        Me.btnPick4.Width = SeperatorLine.ActualWidth

        If gWhoIsInControl = WhoIsInControlType.Main Then
            If gWindowDocked Then
                If (gmwHeight <> Me.ActualHeight) OrElse (Me.Top <> gmwTop) Then
                    Me.Height = gmwHeight
                    Me.Top = gmwTop
                End If
            End If
        Else
            If gWindowDocked Then
                If (gmwHeight <> Me.ActualHeight) OrElse (Me.Top <> gmwTop) Then
                    gmwHeight = Me.ActualHeight
                    gmwWidth = Me.ActualWidth
                    gmwTop = Me.Top
                    gMainWindow.SafelyResizeMainWindow()
                End If

                If (gmwWidth <> Me.ActualWidth) Then ' orelse (Me.Left <> gmwLeft) Then
                    'gmwHeight = Me.ActualHeight
                    'gmwWidth = Me.ActualWidth
                    'gmwTop = Me.Top
                    RezizeMainWindowFromPickAFolder()
                    gMainWindow.SafelyResizeMainWindow()
                End If

            End If
        End If

        If gOverridePickAWindowHeight Then
            gOverridePickAWindowHeight = False
            Me.Height = gmwHeight
        End If

        gWhoIsInControl = WhoIsInControlType.PickAFolder

    End Sub

    Private Sub PickAFolder_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged

        RezizeMainWindowFromPickAFolder()

    End Sub

    Private Sub RezizeMainWindowFromPickAFolder()

        If gWindowDocked Then

            If gWhoIsInControl = WhoIsInControlType.PickAFolder Then

                If My.Settings.DockLeft Then
                    PAFWSaysMWLeftShouldBe = Me.Left + Me.ActualWidth
                Else
                    PAFWSaysMWLeftShouldBe = Me.Left - gmwWidth
                End If

                PAFWSaysMWTopShouldBe = Me.Top

                'need to realign windows, but only if top changes 
                If (gmwTop <> PAFWSaysMWTopShouldBe) Then
                    gMainWindow.SafelyMoveMainWindow()
                End If

            End If

        End If

    End Sub

    Private Sub PickAFolder_StateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.StateChanged

        Try

            If Me.WindowState <> Windows.WindowState.Minimized Then
                If QuickFilterWord.Length = 0 Then
                    LoadTreeView("None")
                Else
                    LoadTreeView(QuickFilterWord)
                End If
            End If

        Catch ex As Exception

        End Try

    End Sub

End Class

