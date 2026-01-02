Imports System.Linq
Imports System.Runtime.InteropServices.ComTypes
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Threading
Imports Microsoft.Office
Imports Microsoft.Office.Interop

Class MainWindow
    Inherits Window

    Private gForegroundColourAlert As System.Windows.Media.SolidColorBrush
    Private gForegroundColourEnabled As System.Windows.Media.SolidColorBrush
    Private gForegroundColourDisabled As System.Windows.Media.SolidColorBrush

    Private gThresholdForReportingProgressOnTheProgressBar As Double

    Private gSuppressUpdatesToDetailBox As Boolean = False

    Enum ActionType As Integer
        Hide = 1
        Delete = 2
        File = 3
    End Enum
    Structure StructureOfActionLog
        Dim FixedIndex As Int16
        Dim ActionApplied As ActionType
        Dim EmailID As String
        Dim SourceStoreID As String
        Dim TargetEntryID As String
        Dim TargetStoreID As String
    End Structure

    Dim ActionLogIndex As Integer = 0
    Dim ActionLogSubIndex As Integer = 0
    Dim ActionLogMaxEntries As Integer = 750
    Dim ActionLogMaxSubEntries As Integer = 1000
    Dim ActionLog(1, 1) As StructureOfActionLog

    Private Enum SortOrder As Integer
        None = 0
        Ascending = 1
        Descending = 2
    End Enum
    Structure StructureOfEmailDetails
        Dim sSubject As String
        Dim sTrailer As String
        Dim sDateAndTime As DateTime
        Dim sTo As String
        Dim sFrom As String
        Dim sCCAs As String
        Dim sOriginalFolderReferenceNumber As Int16
        Dim sRecommendedFolder1ReferenceNumber As Int16
        Dim sRecommendedFolder2ReferenceNumber As Int16
        Dim sRecommendedFolder3ReferenceNumber As Int16
        Dim sRecommendedFolderFinalReferenceNumber As Int16
        Dim sOutlookEntryID As String
        Dim sUnRead As FontWeight
        Dim sMailBoxName As String ' mailbox/postbox name
        Dim sBody As String ' only used when we know we’ll need trailers
    End Structure

    Friend gEmailTable(1) As StructureOfEmailDetails
    Private gEmailTableIndex As Integer = 0
    Private Const gEmailTableGrowth As Integer = 200 ' when more space is needed, grow the table by this many entries

    Private lTotalEMails As Integer = 0
    Private lTotalEMailsToBeReviewed As Integer = 0
    Private lTotalRecommendations As Integer = 0

    Private UniqueSubjectsMap As New Dictionary(Of String, Dictionary(Of Integer, Int16))(StringComparer.Ordinal)

    Private gOriginalWidthSubject, gOriginalWidthTo, gOriginalWidthFrom, gOriginalWidthDate As Integer

    Private gViewSent As Boolean = True
    Private gViewInbox As Boolean = True
    Private gViewAll As Boolean = True
    Private gViewRead As Boolean = True
    Private gViewUnRead As Boolean = True

    Private gFinalRecommendationTable(1) As ListViewRowClass

    Private Shared gIsRefreshing As Boolean = False
    Private ReadOnly gRefreshGateLock As New Object

    Private gCancelRefresh As Boolean = False

    Private Enum QueuedEmailEventType
        Added = 0
        Removed = 1
        Changed = 2
    End Enum

    Private Structure QueuedEmailEvent
        Friend EventType As QueuedEmailEventType
        Friend FolderIndex As Integer
        Friend EntryId As String
        Friend Subject As String
        Friend ToAddr As String
        Friend FromAddr As String
        Friend ReceivedTime As Date
        Friend IsUnread As Boolean
        Friend Body As String
        Friend Attempt As Integer
        Friend MailItem As Microsoft.Office.Interop.Outlook.MailItem
        Friend Folder As Microsoft.Office.Interop.Outlook.MAPIFolder
    End Structure

    Private ReadOnly gListViewEntryIdsLock As New Object
    Private gListViewEntryIdsByFolder As New Dictionary(Of Integer, HashSet(Of String))(IntegerComparer.Instance)

    ' Track Inbox/Sent folders across all mailboxes
    Private gInboxFolderIndices As New List(Of Integer)
    Private gSentFolderIndices As New List(Of Integer)

    ' Per‑store delete target (Deleted Items or Trash) for each Outlook store
    Private Structure StoreDeleteFolderInfo
        Friend StoreId As String
        Friend FolderIndex As Integer
    End Structure

    Private gStoreDeleteFolders As New Dictionary(Of String, StoreDeleteFolderInfo)(StringComparer.OrdinalIgnoreCase)

    ' Number of distinct Outlook mailboxes (stores) detected
    Private Shared _TotalMailBoxes As Integer = 0 ' number of mailboxes in Outlook PostOffice

    Private gOriginalTabControl2Height As Integer

    Private oApp As Microsoft.Office.Interop.Outlook.Application
    Private oNS As Microsoft.Office.Interop.Outlook.NameSpace
    Private oMailItem As Microsoft.Office.Interop.Outlook.MailItem
    Private oTargetFolder As Microsoft.Office.Interop.Outlook.MAPIFolder

    Private gOutlookQuitHooked As Boolean = False

    ' MD5 instance used by ComputeTrailerHash
    Private md5Obj As New System.Security.Cryptography.MD5CryptoServiceProvider

    Private gCurrentlySelectedListViewItemIndex As Integer = 0
    Private Enum SelectionRestoreReason
        Refresh = 0
        Sort = 1
        UserDelete = 2
        OutlookDelete = 3
    End Enum

    Private Class SelectionEntry
        Friend OutlookEntryId As String
        Friend ChainKey As String
        Friend Index As Integer
    End Class

    Private Class SelectionSnapshot
        Friend Entries As List(Of SelectionEntry)
        Friend FirstIndex As Integer
        Friend HasSelection As Boolean
    End Class

    Private gPendingSelectionSnapshot As SelectionSnapshot
    Private gPendingSelectionReason As SelectionRestoreReason = SelectionRestoreReason.Refresh
    Private gPendingSelectionFallbackToFirst As Boolean = True
    Private gPendingSelectionApplied As Boolean = False

    Private lProgressBareRefreshingThresholdCounter As Double = 0

    ' the following weightings are used to level out progress bar increments during a refresh
    ' the current values were determined by measuring actual refreshes of a relatively large set of folders/emails
    ' processing one email takes much longer in the Email Review stage than in the Final stage
    ' the ratio is approximately 30:1
    Private Const lProgressBarWeightingForEmailReviews As Double = 30
    Private Const lProgressBarWeightingForFinalSteps As Double = 1

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.        MainWindow.Visibility = Windows.Visibility.Visible
        EnsureOnlyOneInstanceOfApp()

        gMainWindow = Me

        SetProcessPriorities("Initialize")

    End Sub

    Private gClosingNow As Boolean = False
    Private Sub SafelyCloseWindow()
        Call Dispatcher.BeginInvoke(CloseWindow)
    End Sub
    Private CloseWindow As New System.Windows.Forms.MethodInvoker(AddressOf CloseWindowNow)
    Private Sub CloseWindowNow()

        Me.Dispatcher.BeginInvoke(Sub()
                                      MainWindow.Visibility = Windows.Visibility.Hidden
                                  End Sub)

        Me.Close()
    End Sub


    Private Sub MainWindow_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        If gClosingNow Then Exit Sub

        If gPickAFolderWindow IsNot Nothing Then
            gPickAFolderWindow.Focus()
            gPickAFolderWindow.BringIntoView()
        End If

        Me.Focus()

    End Sub

    Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

#If DEBUG Then
        Console.WriteLine("******************************************************************************************")
#End If

        Try
            'Ensure Settings are kept thru an upgrade
            MainWindow.Visibility = Windows.Visibility.Hidden

            ' Verify Outlook is available 
            ' note: if Outlook is not running and the user selects not to start it, then FileFriendly will exit immediately
            If Not EnsureOutlookIsRunning() Then
                Exit Sub
            End If

            MainWindow.Visibility = Windows.Visibility.Visible

            Try
                Dim version As String = oApp.Version
                If String.IsNullOrEmpty(version) Then
                    Throw New Exception("Outlook version is empty")
                End If
            Catch

                Call ShowMessageBox("FileFriendly - Critical Error",
                     CustomDialog.CustomDialogIcons.Stop,
                     "It appears that Microsoft Outlook is not installed or accessible on this computer.",
                     "FileFriendly requires Outlook to be able to run.",
                     "MainWindow_Loaded",
                     "",
                     CustomDialog.CustomDialogIcons.None,
                     CustomDialog.CustomDialogButtons.OK,
                     CustomDialog.CustomDialogResults.OK)

                End
            End Try

            gForegroundColourAlert = System.Windows.Media.Brushes.Red
            gForegroundColourEnabled = Me.MenuAbout.Foreground
            gForegroundColourDisabled = System.Windows.Media.Brushes.Gray

            Dim a As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()

            gAppVersion = a.GetName().Version
            gAppVersionString = gAppVersion.ToString

            Dim webFriendlyVersionNumber As String = gAppVersionString.Replace(".", "_")
            While webFriendlyVersionNumber.EndsWith("_0")
                webFriendlyVersionNumber = webFriendlyVersionNumber.Remove(webFriendlyVersionNumber.Length - 2)
            End While

            gHelpWebPage &= "Help_v" & webFriendlyVersionNumber & ".md"

            ReDim gDockSound(My.Resources.dock.Length)

            My.Resources.dock.Read(gDockSound, 0, My.Resources.dock.Length)

            If My.Settings.ApplicationVersion <> gAppVersion.ToString Then
                My.Settings.Upgrade()
                My.Settings.ApplicationVersion = gAppVersionString
            End If

            Me.ListView1.Visibility = Windows.Visibility.Hidden

            Dim j As String = "  ".Trim
            Dim i As Integer = j.Length

            'set screen width and height (managing the case where the client has changed screen resolutions from last run)

            '****** width
            Dim dCurrentScreenWidth As Double = System.Windows.SystemParameters.PrimaryScreenWidth
            If My.Settings.ScreenWidth = dCurrentScreenWidth Then
                ' no need to change settings from last time
            Else
                My.Settings.ScreenWidth = dCurrentScreenWidth

                Dim CombindedWindowWidth = My.Settings.MainWidth + My.Settings.FoldersWidth

                Select Case dCurrentScreenWidth
                    Case Is > CombindedWindowWidth
                        ' no adjustments necessary

                    Case Is <= 800
                        My.Settings.FoldersWidth = 400
                        My.Settings.MainWidth = 750
                        My.Settings.StartDocked = False
                        My.Settings.MainLeft = 50
                        My.Settings.FoldersLeft = 0

                    Case Is <= 1152
                        My.Settings.FoldersWidth = 500
                        My.Settings.MainWidth = 800
                        My.Settings.StartDocked = False
                        My.Settings.FoldersLeft = 0
                        My.Settings.MainLeft = 50

                    Case Is <= 1280
                        My.Settings.FoldersWidth = 450
                        My.Settings.MainWidth = 800
                        My.Settings.StartDocked = True
                        My.Settings.FoldersLeft = 0
                        My.Settings.MainLeft = My.Settings.MainWidth

                    Case Else
                        My.Settings.FoldersWidth = 550
                        My.Settings.MainWidth = 800
                        My.Settings.StartDocked = True
                        My.Settings.FoldersLeft = 0
                        My.Settings.MainLeft = My.Settings.MainWidth

                End Select

            End If

            Me.Width = My.Settings.MainWidth
            Me.Left = My.Settings.MainLeft
            gmwWidth = Me.ActualWidth
            gmwLeft = MainWindow.Left

            '****** height 

            MainWindow.Top = My.Settings.MainTop
            gmwTop = MainWindow.Top
            gmwHeight = Me.ActualHeight

            Dim dCurrentScreenHeight As Double = System.Windows.SystemParameters.PrimaryScreenHeight
            If My.Settings.ScreenHeight = dCurrentScreenHeight Then
                ' no need to change settings from last time
            Else
                Const HieghtOfSystrayBar As Double = 30
                If Me.Height > System.Windows.SystemParameters.PrimaryScreenHeight - HieghtOfSystrayBar Then
                    gmwTop = 0
                    MainWindow.Top = 0
                    My.Settings.MainTop = 0
                    My.Settings.FoldersTop = 0
                    gmwHeight = System.Windows.SystemParameters.PrimaryScreenHeight - HieghtOfSystrayBar
                    Me.Height = gmwHeight
                    gOverridePickAWindowHeight = True
                End If
            End If

            gWindowDocked = My.Settings.StartDocked

            gRefreshInbox = My.Settings.ScanInbox
            gRefreshSent = My.Settings.ScanSent
            gRefreshOtherFolders = My.Settings.ScanAll

            Me.MenuViewRead.IsChecked = gViewRead
            Me.MenuViewUnRead.IsChecked = gViewUnRead

            gAutoChainSelect = My.Settings.AutoChainSelect

            Me.lblMainMessageLine.Content = ""
            gOriginalTabControl2Height = Me.TabControl2.ActualHeight

            Me.Label7.Visibility = Windows.Visibility.Hidden
            Me.Row3.Height = New System.Windows.GridLength(Me.Row3.ActualHeight - 20, GridUnitType.Auto)
            Me.TabControl2.Height = Me.TabControl2.ActualHeight - 20

            MenuOptionEnabled("Undo", False)

            MenuOptionEnabled("Refresh", False)

            CheckIfNewVersionIsAvailable()

            RefreshGrid(True, False, False)

            ReDim ActionLog(ActionLogMaxSubEntries, ActionLogMaxSubEntries)

            My.Settings.Save()

            MemoryManagement.FlushMemory()

            ' start monitoring Outlook for new emails after a short delay to allow the main window to finish loading

            Dim monitoringInitTimer As New DispatcherTimer() With {
             .Interval = TimeSpan.FromMilliseconds(500)
             }

            AddHandler monitoringInitTimer.Tick, Sub()
                                                     monitoringInitTimer.Stop()
                                                     RemoveHandler monitoringInitTimer.Tick, Nothing
                                                     Try
                                                         InitializeMonitoringOfOutlookEvents()
                                                     Catch
                                                     End Try
                                                 End Sub
            monitoringInitTimer.Start()
            Thread.Sleep(500)

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Loading Error",
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

    Private Sub EnsureOnlyOneInstanceOfApp()

        Try

            Dim appProc() As Process
            Dim strModName, strProcName As String
            strModName = Process.GetCurrentProcess.MainModule.ModuleName
            strProcName = System.IO.Path.GetFileNameWithoutExtension(strModName)

            appProc = Process.GetProcessesByName(strProcName)

            If appProc.Length > 1 Then

                Call ShowMessageBox("FileFriendly - Start Warning",
                     CustomDialog.CustomDialogIcons.Warning,
                     "FileFriendly is already running",
                     "Only one instance of FileFriendly can be run at once.",
                     "The original instance of FileFriendly will remain running, but a new one will not be started",
                     "",
                     CustomDialog.CustomDialogIcons.None,
                     CustomDialog.CustomDialogButtons.OK,
                     CustomDialog.CustomDialogResults.OK)

                End

            End If

        Catch ex As Exception
        End Try

    End Sub

    Private Sub GracefulShutdown()

        On Error Resume Next

        SetProcessPriorities("Shutdown")

        My.Settings.MainWidth = Me.ActualWidth
        My.Settings.MainTop = Me.Top
        My.Settings.MainLeft = Me.Left

        My.Settings.FoldersWidth = gPickAFolderWindow.ActualWidth
        My.Settings.FoldersTop = gPickAFolderWindow.Top
        My.Settings.FoldersLeft = gPickAFolderWindow.Left

        'this should always be true, but check anyway
        If System.Windows.SystemParameters.PrimaryScreenWidth > 0 Then
            My.Settings.ScreenWidth = System.Windows.SystemParameters.PrimaryScreenWidth
        End If

        My.Settings.Save()

        If oMailItem IsNot Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oMailItem)
            oMailItem = Nothing
        End If

        If oTargetFolder IsNot Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oTargetFolder)
            oTargetFolder = Nothing
        End If

        If oNS IsNot Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oNS)
            oNS = Nothing
        End If

        If oApp IsNot Nothing Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oApp)
            oApp = Nothing
        End If

        System.GC.Collect()
        System.GC.WaitForPendingFinalizers()

    End Sub

    Private Sub MainWindow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing

        If Not gClosingNow Then
            gClosingNow = True
            Me.Visibility = Windows.Visibility.Hidden

            If gPickAFolderWindow IsNot Nothing Then
                gPickAFolderWindow.Visibility = Windows.Visibility.Hidden
            End If
        End If

        ClearMonitoringOfOutlookEvents()

        GracefulShutdown()

    End Sub

    Public Sub SafelyMoveMainWindow()
        Call Dispatcher.BeginInvoke(MoveMainWindow)
    End Sub
    Private MoveMainWindow As New System.Windows.Forms.MethodInvoker(AddressOf MoveMainWindowNow)
    Private Sub MoveMainWindowNow()

        PlaceWindow()

    End Sub

    Private Sub PlaceWindow()

        gWhoIsInControl = WhoIsInControlType.PickAFolder
        If Me.Top <> PAFWSaysMWTopShouldBe Then Me.Top = PAFWSaysMWTopShouldBe
        If Me.Left <> PAFWSaysMWLeftShouldBe Then Me.Left = PAFWSaysMWLeftShouldBe

    End Sub

    Private Sub MainWindow_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles Me.MouseEnter
        Try
            ' Find any open PickAFolder window
            Dim pick As PickAFolder = Nothing

            For Each w As System.Windows.Window In System.Windows.Application.Current.Windows
                pick = TryCast(w, PickAFolder)
                If pick IsNot Nothing Then
                    Exit For
                End If
            Next

            ' If PickAFolder is active and mouse enters main window, activate main window
            If pick IsNot Nothing AndAlso pick.IsActive AndAlso Not Me.IsActive Then
                Me.Activate()
            End If
        Catch ex As Exception
            ' Swallow or log as per your existing style
        End Try
    End Sub

    Public Sub SafelyUpdateContextMenu()
        Call Dispatcher.BeginInvoke(UpdateContextMenu)
    End Sub
    Private UpdateContextMenu As New System.Windows.Forms.MethodInvoker(AddressOf UpdateContextMenuNow)
    Private Sub UpdateContextMenuNow()

        Try

            Me.ContextMenuSeperator.Visibility = Windows.Visibility.Collapsed

            If gContextFile1.Length > 0 Then
                Me.MenuContextFile1.Header = "_File in " & gContextFile1
                Me.MenuContextFile1.Visibility = Windows.Visibility.Visible
                Me.ContextMenuSeperator.Visibility = Windows.Visibility.Visible
            Else
                Me.MenuContextFile1.Visibility = Windows.Visibility.Collapsed
            End If

            If gContextFile2.Length > 0 Then
                Me.MenuContextFile2.Header = "F_ile in " & gContextFile2
                Me.MenuContextFile2.Visibility = Windows.Visibility.Visible
                Me.ContextMenuSeperator.Visibility = Windows.Visibility.Visible
            Else
                Me.MenuContextFile2.Visibility = Windows.Visibility.Collapsed
            End If

            If gContextFile3.Length > 0 Then
                Me.MenuContextFile3.Header = "Fi_le in " & gContextFile3
                Me.MenuContextFile3.Visibility = Windows.Visibility.Visible
                Me.ContextMenuSeperator.Visibility = Windows.Visibility.Visible
            Else
                Me.MenuContextFile3.Visibility = Windows.Visibility.Collapsed
            End If

            If gContextFile4.Length > 0 Then
                Me.MenuContextFile4.Header = "Fil_e in " & gContextFile4
                Me.MenuContextFile4.Visibility = Windows.Visibility.Visible
                Me.ContextMenuSeperator.Visibility = Windows.Visibility.Visible
            Else
                Me.MenuContextFile4.Visibility = Windows.Visibility.Collapsed
            End If

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly",
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
    Public Sub SafelyResizeMainWindow()
        Call Dispatcher.BeginInvoke(ResizeMainWindow)
    End Sub
    Private ResizeMainWindow As New System.Windows.Forms.MethodInvoker(AddressOf ResizeMainWindowNow)
    Private Sub ResizeMainWindowNow()

        If Me.Top <> gmwTop Then
            Me.Top = gmwTop
        End If

        If Me.ActualHeight <> gmwHeight Then
            Me.Height = gmwHeight
        End If

        If gWindowDocked Then
            PlaceWindow()
        End If

    End Sub

    Public intRecommendationFinal As String = ""
    Public Sub SafelyUpdateRecommendationFromPickAFolderWindow()
        Call Dispatcher.BeginInvoke(UpdateRecommendation)
    End Sub
    Private UpdateRecommendation As New System.Windows.Forms.MethodInvoker(AddressOf UpdateRecommendationNow)
    Private Sub UpdateRecommendationNow()

        PerformAction("File", False)
        gPickFromContextMenuOverride = -1

    End Sub

    Private Sub SetProcessPriorities(ByVal Command As String)

        Static Dim myProcess, OutlookProcess As Process

        Try

            Select Case Command

                Case Is = "Initialize"

                    ' Do not spin up Outlook here; let EnsureOutlookIsRunning()
                    ' manage starting and repairing the Outlook session as needed.
                    oApp = Nothing
                    oNS = Nothing

                    Dim mySessionID As Int16
                    Dim AllProcesses() As System.Diagnostics.Process = Process.GetProcesses()

                    myProcess = Process.GetCurrentProcess
                    mySessionID = myProcess.SessionId

                    For Each Process In AllProcesses
                        If Process.SessionId = mySessionID Then
                            If Process.ProcessName.ToUpper.StartsWith("OUTLOOK") Then
                                OutlookProcess = Process
                                Exit For
                            End If
                        End If
                    Next

                    AllProcesses = Nothing

                Case Is = "Start Outlook Review"

                    myProcess.PriorityClass = ProcessPriorityClass.AboveNormal
                    If OutlookProcess IsNot Nothing Then OutlookProcess.PriorityClass = ProcessPriorityClass.AboveNormal


                Case Is = "End Outlook Review"

                    If OutlookProcess IsNot Nothing Then OutlookProcess.PriorityClass = ProcessPriorityClass.Normal

                Case Is = "End Review"

                    myProcess.PriorityClass = ProcessPriorityClass.Normal

                Case Is = "Shutdown"

                    If OutlookProcess IsNot Nothing Then OutlookProcess.PriorityClass = ProcessPriorityClass.Normal

                    If OutlookProcess IsNot Nothing Then OutlookProcess.Dispose()

                    myProcess.Dispose()

            End Select

        Catch ex As Exception

        End Try

    End Sub

#Region "List View Stuff"

    ' Thread‑safe wrapper to update the cursor from any thread
    Private Sub SetMousePointer(ByVal cursor As System.Windows.Input.Cursor)
        If Dispatcher.CheckAccess() Then
            Me.Cursor = cursor
        Else
            Dispatcher.BeginInvoke(Sub() Me.Cursor = cursor)
        End If
    End Sub

    ' Keeps the Wait cursor active across scheduled + running refresh work.
    Private ReadOnly gRefreshCursorLock As New Object()
    Private gRefreshCursorRefCount As Integer = 0

    Private Sub BeginRefreshCursor()
        Dim shouldSetWait As Boolean = False

        SyncLock gRefreshCursorLock
            gRefreshCursorRefCount += 1
            If gRefreshCursorRefCount = 1 Then
                shouldSetWait = True
            End If
        End SyncLock

        If shouldSetWait Then
            SetMousePointer(Cursors.Wait)
        End If
    End Sub

    Private Sub EndRefreshCursor()
        Dim shouldSetArrow As Boolean = False

        SyncLock gRefreshCursorLock
            If gRefreshCursorRefCount > 0 Then
                gRefreshCursorRefCount -= 1
                If gRefreshCursorRefCount = 0 Then
                    shouldSetArrow = True
                End If
            End If
        End SyncLock

        If shouldSetArrow Then
            SetMousePointer(Cursors.Arrow)
        End If
    End Sub


    Delegate Sub ShowFoldersCallback()
    Private Sub ShowFolders()

        If Me.WindowState <> Windows.WindowState.Minimized Then
            ShowFolderWindow()
        End If

    End Sub

    Delegate Sub BeginLoadCallback()
    Private Sub BeginLoad()

        MenuOptionEnabled("Open", False)
        MenuOptionEnabled("Hide", False)
        MenuOptionEnabled("Delete", False)
        MenuOptionEnabled("Undo", False)
        MenuOptionEnabled("Refresh", False)
        MenuOptionEnabled("Options", False)
        MenuOptionEnabled("View", False)

        gIsRefreshing = True
        gCancelRefresh = False

        UpdateRefreshMenuState()

        gViewInbox = gRefreshInbox
        gViewSent = gRefreshSent
        gViewAll = gRefreshOtherFolders

        MenuViewInbox.IsChecked = gRefreshInbox
        MenuViewSent.IsChecked = gRefreshSent
        MenuViewAll.IsChecked = gRefreshOtherFolders

        If gViewAll Or gViewInbox Or gViewSent Then
            MenuView.Foreground = gForegroundColourEnabled
            gViewRead = True
            gViewUnRead = True
            MenuViewRead.IsChecked = True
            MenuViewUnRead.IsChecked = True
        Else
            MenuView.Foreground = gForegroundColourDisabled
            gViewRead = False
            gViewUnRead = False
            MenuViewRead.IsChecked = False
            MenuViewUnRead.IsChecked = False
        End If

    End Sub

    Delegate Sub FinalizeLoadCallback(ByVal MSOutlookDrivenEvent As Boolean)
    Private Sub FinalizeLoad(ByVal MSOutlookDrivenEvent As Boolean)

        ApplyFilter()

        ' Adjust Mailbox column after items are loaded
        UpdateMailboxColumnVisibility()
        RecalculateListViewColumnWidths()

        gIsRefreshing = False
        gCancelRefresh = False
        UpdateRefreshMenuState()

        MenuOptionEnabled("Options", True)

        If ListView1.Items.Count > 0 Then

            Me.MenuActions.Foreground = gForegroundColourEnabled

            MenuOptionEnabled("Open", True)
            MenuOptionEnabled("Hide", True)
            MenuOptionEnabled("Delete", True)
            If ActionLogIndex > 0 Then MenuOptionEnabled("Undo", True)
            MenuOptionEnabled("View", True)

            Me.ListView1.Focus()

        Else

            If gRefreshOtherFolders Or gRefreshSent Or gRefreshInbox Then
                Me.MenuRefresh.Foreground = gForegroundColourEnabled
                Me.MenuActions.Foreground = gForegroundColourEnabled
            Else
                Me.MenuRefresh.Foreground = gForegroundColourAlert
                Me.MenuActions.Foreground = gForegroundColourAlert
            End If

        End If

        ' Play a beep if that option is set in the settings except:
        ' if this was a MS Outlook driven event (as opposed to a startup or user initiated refresh) 
        If My.Settings.SoundScanComplete Then

            If (Not MSOutlookDrivenEvent) Then
                Beep()
            End If

        End If

        If gRefreshQueued Then
            gRefreshQueued = False
            ScheduleRefreshGrid()
        End If

        ' Mark scheduled refresh complete and release the cursor *after* any queued refresh decision above.
        SyncLock gRefreshGridLock
            gRefreshGridScheduled = False
        End SyncLock

        EndRefreshCursor()

    End Sub

    Delegate Sub SetFolderNameTextCallback(ByVal [text] As String)
    Private Sub SetFoldersNameText(ByVal [text] As String)
        Me.lblMainMessageLine.Content = [text]
    End Sub

    Delegate Sub SetToolTipCallback(ByVal [text] As String)
    Private Sub SetToolTip(ByVal [text] As String)

        If [text] = "Done" Then
            Me.lblMainMessageLine.ToolTip = Me.lblMainMessageLine.ToolTip.replace("are being", "were")
        Else
            Me.lblMainMessageLine.ToolTip = [text]
        End If

    End Sub

    Delegate Sub SetProgressBarMaxValueCallback(ByVal [Double] As Double)
    Private Sub SetProgressBarMaxValue(ByVal [Double] As Double)
        Me.ProgressBar1.Maximum = [Double]
    End Sub

    Delegate Sub SetProgressBarValueCallback(ByVal [Double] As Double)
    Private Sub SetProgressBarValue(ByVal [Double] As Double)
        Me.ProgressBar1.Value = [Double]
    End Sub

    Delegate Sub SetProgressBarVisableCallback(ByVal [WindowsVisibility] As Windows.Visibility)
    Private Sub SetProgressBarVisable(ByVal [WindowsVisibility] As Windows.Visibility)
        Me.ProgressBar1.Visibility = WindowsVisibility
    End Sub

    Public Class ListViewRowClass

        Public Enum ChainIndicatorValues As Integer
            NotPartOfAChain = 0
            TopOfTheChain = 1
            MiddleOfTheChain = 2
            EndOfTheChain = 3
        End Enum

        Private _MailBoxName As String
        Public Property MailBoxName() As String
            Get
                Return Me._MailBoxName
            End Get
            Set(ByVal value As String)
                Me._MailBoxName = value
            End Set
        End Property

        Private _Index As Integer
        Public Property Index() As Integer
            Get
                Return Me._Index
            End Get
            Set(ByVal value As Integer)
                Me._Index = value
            End Set
        End Property

        Private _FixedIndex As Integer
        Public Property FixedIndex() As Integer
            Get
                Return Me._FixedIndex
            End Get
            Set(ByVal value As Integer)
                Me._FixedIndex = value
            End Set
        End Property

        Private _ChainIndicator As Integer
        Public Property ChainIndicator() As ChainIndicatorValues
            ' 0 not part of an email chain
            ' 1 most recent email of a chain
            ' 2 middle of an email chain
            ' 3 original email of chain
            Get
                Return Me._ChainIndicator
            End Get
            Set(ByVal value As ChainIndicatorValues)
                Me._ChainIndicator = value
            End Set
        End Property

        Private _Subject As String
        Public Property Subject() As String
            Get
                Return Me._Subject
            End Get
            Set(ByVal value As String)
                Me._Subject = value
            End Set
        End Property

        Private _Trailer As String
        Public Property Trailer() As String
            Get
                Return Me._Trailer
            End Get
            Set(ByVal value As String)
                Me._Trailer = value
            End Set
        End Property

        Private _From As String
        Public Property From() As String
            Get
                Return Me._From
            End Get
            Set(ByVal value As String)
                Me._From = value
            End Set
        End Property

        Private _xTo As String
        Public Property xTo() As String
            Get
                Return Me._xTo
            End Get
            Set(ByVal value As String)
                Me._xTo = value
            End Set
        End Property

        Private _DateTime As String
        Public Property DateTime() As Date
            Get
                Return Me._DateTime
            End Get
            Set(ByVal value As Date)
                Me._DateTime = value
            End Set
        End Property

        Private _OriginalFolder As Integer
        Public Property OriginalFolder() As Integer
            Get
                Return Me._OriginalFolder
            End Get
            Set(ByVal value As Integer)
                Me._OriginalFolder = value
            End Set
        End Property

        Private _RecommendedFolderFinal As Integer
        Public Property RecommendedFolderFinal() As Integer
            Get
                Return Me._RecommendedFolderFinal
            End Get
            Set(ByVal value As Integer)
                Me._RecommendedFolderFinal = value
            End Set
        End Property

        Private _RecommendedFolder1 As Integer
        Public Property RecommendedFolder1() As Integer
            Get
                Return Me._RecommendedFolder1
            End Get
            Set(ByVal value As Integer)
                Me._RecommendedFolder1 = value
            End Set
        End Property

        Private _RecommendedFolder2 As Integer
        Public Property RecommendedFolder2() As Integer
            Get
                Return Me._RecommendedFolder2
            End Get
            Set(ByVal value As Integer)
                Me._RecommendedFolder2 = value
            End Set
        End Property

        Private _RecommendedFolder3 As Integer
        Public Property RecommendedFolder3() As Integer
            Get
                Return Me._RecommendedFolder3
            End Get
            Set(ByVal value As Integer)
                Me._RecommendedFolder3 = value
            End Set
        End Property

        Private _OutlookEntryID As String
        Public Property OutlookEntryID() As String
            Get
                Return Me._OutlookEntryID
            End Get
            Set(ByVal value As String)
                Me._OutlookEntryID = value
            End Set
        End Property

        Private _Unread As FontWeight
        Public Property UnRead() As FontWeight
            Get
                Return Me._Unread
            End Get
            Set(ByVal value As FontWeight)
                Me._Unread = value
            End Set
        End Property

    End Class

#End Region

    Private Sub UpdateRefreshMenuState()
        If gIsRefreshing Then
            Me.MenuRefresh.Header = "_Stop Refresh"
            Me.MenuRefresh.IsEnabled = True
            Me.MenuRefresh.Foreground = gForegroundColourEnabled
        Else
            Me.MenuRefresh.Header = "_Refresh"
            ' Enabled/disabled and colour based on current rules
            MenuOptionEnabled("Refresh", (gRefreshInbox Or gRefreshSent Or gRefreshOtherFolders))
        End If
    End Sub


    Delegate Sub ClearListView1Callback()
    Private Sub ClearListView1()

        ListView1.Items.Clear()

    End Sub

    Delegate Sub SetListViewItemCallback(ByVal FinalRecommendationTable As ListViewRowClass())

    Private Sub SetListViewItem(ByVal FinalRecommendationTable As ListViewRowClass())

        Try

            If FinalRecommendationTable.Length = 0 Then
                Return
            End If

            Dim lCurrentSubjectPlusTrailer As String = ""
            Dim lPreviousSubjectPlusTrailer As String = ""

            Dim LineCount As Integer = 0

            ListView1.Items.Clear()

            ResetEMailChainRelationShips(FinalRecommendationTable)

            ' print the contents of the FinalRecommendationTable to the console for debugging
            'For x As Integer = 0 To FinalRecommendationTable.Length - 1
            '    Dim row = FinalRecommendationTable(x)
            '    If row IsNot Nothing Then
            '        Console.WriteLine($"Row {x}: Subject = {row.Subject}, Trailer = {row.Trailer}")
            '    End If
            'Next

            If gCurrentSortOrder = "Subject" Then
                SetEMailChainRelationShips(FinalRecommendationTable)
            End If

            For x As Integer = 0 To FinalRecommendationTable.Length - 1

                Dim row = FinalRecommendationTable(x)
                If row Is Nothing Then Continue For

                If row.Index = -1 Then
                    gFinalRecommendationTable(x).Index = -1
                Else
                    row.Index = LineCount
                    gFinalRecommendationTable(x).FixedIndex = LineCount
                    ListView1.Items.Add(row)
                End If

                LineCount += 1

            Next

            Me.ListView1.Visibility = Windows.Visibility.Visible

            Static Dim ivebeenbumped As Boolean = False

            Try
                'Bump the window once so the silly thing aligns
                If ivebeenbumped Then
                Else
                    ivebeenbumped = True
                    Me.Width -= 1
                    Me.Width += 1
                End If
            Catch ex As Exception
            End Try

            RecalculateListViewColumnWidths()

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Set ListView Item Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 ex.TargetSite.Name & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try

    End Sub

    Private Sub ResetEMailChainRelationShips(ByRef FinalRecommendationTable As ListViewRowClass())

        If FinalRecommendationTable Is Nothing Then Exit Sub

        For x As Integer = 0 To FinalRecommendationTable.Length - 1
            Dim row = FinalRecommendationTable(x)
            If row Is Nothing Then Continue For
            row.ChainIndicator = ListViewRowClass.ChainIndicatorValues.NotPartOfAChain
        Next

    End Sub

    Private Sub SetEMailChainRelationShips(ByRef FinalRecommendationTable As ListViewRowClass())

        If FinalRecommendationTable Is Nothing Then Exit Sub

        'Set middles of chains
        For x As Integer = 1 To FinalRecommendationTable.Length - 1
            Dim prev = FinalRecommendationTable(x - 1)
            Dim cur = FinalRecommendationTable(x)
            If prev Is Nothing OrElse cur Is Nothing Then Continue For
            If prev.Subject = cur.Subject AndAlso prev.Trailer = cur.Trailer Then
                cur.ChainIndicator = ListViewRowClass.ChainIndicatorValues.MiddleOfTheChain
            End If
        Next

        'Set tops
        For x As Integer = 0 To FinalRecommendationTable.Length - 2
            Dim cur = FinalRecommendationTable(x)
            Dim nxt = FinalRecommendationTable(x + 1)
            If cur Is Nothing OrElse nxt Is Nothing Then Continue For
            If cur.ChainIndicator = ListViewRowClass.ChainIndicatorValues.NotPartOfAChain AndAlso
           nxt.ChainIndicator = ListViewRowClass.ChainIndicatorValues.MiddleOfTheChain Then
                cur.ChainIndicator = ListViewRowClass.ChainIndicatorValues.TopOfTheChain
            End If
        Next

        'Set bottoms
        For x As Integer = 1 To FinalRecommendationTable.Length - 2
            Dim cur = FinalRecommendationTable(x)
            Dim nxt = FinalRecommendationTable(x + 1)
            If cur Is Nothing OrElse nxt Is Nothing Then Continue For
            If cur.ChainIndicator = ListViewRowClass.ChainIndicatorValues.MiddleOfTheChain AndAlso
           nxt.ChainIndicator <> ListViewRowClass.ChainIndicatorValues.MiddleOfTheChain Then
                cur.ChainIndicator = ListViewRowClass.ChainIndicatorValues.EndOfTheChain
            End If
        Next

        'special case to deal with final entryId
        If FinalRecommendationTable.Length > 1 Then
            Dim last = FinalRecommendationTable(FinalRecommendationTable.Length - 1)
            If last IsNot Nothing AndAlso
           last.ChainIndicator = ListViewRowClass.ChainIndicatorValues.MiddleOfTheChain Then
                last.ChainIndicator = ListViewRowClass.ChainIndicatorValues.EndOfTheChain
            End If
        End If

    End Sub
    Private Sub ApplyFilter()

        Try

            Dim LineCount As Integer = 0
            Dim InboxItem, SentItem, NeitherInboxNorSentItem, MessageWasRead As Boolean

            ListView1.Items.Clear()

            ' If nothing has been loaded into the recommendation table, just show 0 e-mails
            If gFinalRecommendationTable Is Nothing OrElse gFinalRecommendationTable.Length = 0 Then
                Me.lblMainMessageLine.Content = "0 e-mails"
                Exit Try
            End If

            If Not (gRefreshInbox Or gRefreshSent Or gRefreshOtherFolders) Then
                Me.lblMainMessageLine.Content = "0 e-mails"
                Exit Try
            End If

            Dim NewRecommendationTable(gFinalRecommendationTable.Length - 1) As ListViewRowClass

            For x As Integer = 0 To gFinalRecommendationTable.Length - 1

                Dim row = gFinalRecommendationTable(x)
                If row Is Nothing Then Continue For

                MessageWasRead = (row.UnRead = System.Windows.FontWeights.Normal)

                If (gViewRead And MessageWasRead) Or (gViewUnRead And Not MessageWasRead) Then

                    InboxItem = gInboxFolderIndices.Contains(row.OriginalFolder)
                    SentItem = gSentFolderIndices.Contains(row.OriginalFolder)
                    NeitherInboxNorSentItem = Not (InboxItem Or SentItem)

                    If (gViewInbox And InboxItem) Or
                       (gViewSent And SentItem) Or
                       (gViewAll And NeitherInboxNorSentItem) Then

                        If row.Index <> -1 Then
                            row.Index = LineCount
                            NewRecommendationTable(LineCount) = row
                            LineCount += 1
                        End If

                    End If
                End If

            Next

            ' No items passed the filter
            If LineCount = 0 Then
                Me.lblMainMessageLine.Content = "0 e-mails"
                UpdateMainMessageLine()
                Exit Try
            End If

            ' Trim the array to the actual count
            ReDim Preserve NewRecommendationTable(LineCount - 1)

            ResetEMailChainRelationShips(NewRecommendationTable)

            If gCurrentSortOrder = "Subject" Then
                SetEMailChainRelationShips(NewRecommendationTable)
            End If

            For x As Integer = 0 To NewRecommendationTable.Length - 1
                If NewRecommendationTable(x) IsNot Nothing Then
                    ListView1.Items.Add(NewRecommendationTable(x))
                End If
            Next

            NewRecommendationTable = Nothing

            If ListView1.Items.Count > 0 Then
                Me.ListView1.Focus()
            End If

            RestorePendingSelection()

            UpdateSortHeaderGlyph()

        Catch ex As Exception
            ' Optional: log ex.ToString()
        End Try

    End Sub

    Delegate Sub SetListViewSelectedItemCallback()
    Private Sub SetListViewSelectedItem()

        ' set the listview selected item based on gCurrentlySelectedListViewItemIndex

        Try
            ListView1.SelectedIndex = gCurrentlySelectedListViewItemIndex
        Catch
            Try
                ListView1.SelectedIndex = 0
            Catch
                ListView1.SelectedIndex = -1
            End Try
        End Try

    End Sub

    Private Sub ClearGrid()

        Me.ListView1.Visibility = Windows.Visibility.Hidden

        ' Remove arrow from previously sorted header
        If _lastheaderClicked IsNot Nothing Then
            _lastheaderClicked.Column.HeaderTemplate = Nothing
        End If

        'update menu bar 
        BeginLoad()
        MenuActions.Foreground = gForegroundColourAlert
        MenuRefresh.Foreground = gForegroundColourAlert
        MenuRefresh.IsEnabled = True

        lblMainMessageLine.Content = "0 e-mails (0 selected)"

        BlankOutDetails()

    End Sub

    Private gRefreshQueued As Boolean = False

    Private Sub RefreshGrid(ByVal InitialLoad As Boolean, ByVal MSOutlookDrivenEvent As Boolean, ByVal QuickRefresh As Boolean)

        Try

            Dim shouldStart As Boolean

            SyncLock gRefreshGateLock
                If gIsRefreshing Then
                    gRefreshQueued = True
                    shouldStart = False
                Else
                    gIsRefreshing = True
                    shouldStart = True
                End If
            End SyncLock

            If Not shouldStart Then
                Return
            End If

            ' Remove arrow from previously sorted header
            If _lastheaderClicked IsNot Nothing Then
                _lastheaderClicked.Column.HeaderTemplate = Nothing
            End If

            BlankOutDetails()

            ' Use the thread pool instead of creating a raw Thread per refresh.
            Task.Run(Sub() RefreshBackGroundTask(InitialLoad, MSOutlookDrivenEvent, QuickRefresh))

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Refresh Grid Error",
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

    Private Collection_of_folders_to_exclude = New System.Collections.Specialized.StringCollection
    Private Collection_of_folders_to_exclude_is_empty As Boolean = True

    Private Sub RefreshBackGroundTask(ByVal InitialLoad As Boolean, ByVal MSOutlookDrivenEvent As Boolean, ByVal QuickRefresh As Boolean)


#If DEBUG Then

        'time how long the overall process takes (when in debug mode)
        Dim swOverall As Stopwatch = Stopwatch.StartNew()
        swOverall.Start()

#End If

        Try

            MemoryManagement.FlushMemory()

            If MSOutlookDrivenEvent OrElse QuickRefresh Then

                ' if this is an Outlook driven event we skip the finding of all folders

            Else

                If InitialLoad OrElse gRefreshOtherFolders Then

                    Me.Dispatcher.BeginInvoke(New BeginLoadCallback(AddressOf BeginLoad), New Object() {})

                    ' if cancelled before we even start, honour it
                    If gCancelRefresh Then GoTo CleanExit

                    SetProcessPriorities("Start Outlook Review")

                    Me.Dispatcher.BeginInvoke(New SetToolTipCallback(AddressOf SetToolTip), New Object() {"Folders are being reviewed"})

                    Collection_of_folders_to_exclude = My.Settings.ExcludedScanFolders 'list of all folders to be excluded from scan

                    Collection_of_folders_to_exclude_is_empty = (Collection_of_folders_to_exclude Is Nothing)

                    Me.Dispatcher.BeginInvoke(New ShowFoldersCallback(AddressOf ShowFolders), New Object() {})

                    FindAllFolders()

                End If

            End If


            gMinimizeMaximizeAllowed = True

            Me.Dispatcher.BeginInvoke(New ShowFoldersCallback(AddressOf ShowFolders), New Object() {})

            If gRefreshInbox OrElse gRefreshSent OrElse gRefreshOtherFolders Then

                If lTotalEMailsToBeReviewed > 0 Then

                    Me.Dispatcher.BeginInvoke(New SetFolderNameTextCallback(AddressOf SetFoldersNameText), New Object() {"Reviewing " & lTotalEMailsToBeReviewed.ToString("#,#", System.Globalization.CultureInfo.InvariantCulture) & " of " & lTotalEMails.ToString("#,#", System.Globalization.CultureInfo.InvariantCulture) & " e-mails"})

                    Me.Dispatcher.BeginInvoke(New SetProgressBarVisableCallback(AddressOf SetProgressBarVisable), New Object() {Windows.Visibility.Visible})

                    ProcessAllFolders(MSOutlookDrivenEvent, QuickRefresh)

                    If Not gCancelRefresh Then

                        Collection_of_folders_to_exclude = Nothing
                        Collection_of_folders_to_exclude_is_empty = True

                        SetProcessPriorities("End Outlook Review")

                        MemoryManagement.FlushMemory()

                        EstablishRecommendations()

                        UpdateListView()

                        Me.Dispatcher.BeginInvoke(New SetProgressBarVisableCallback(AddressOf SetProgressBarVisable), New Object() {Windows.Visibility.Hidden})

                        Me.Dispatcher.BeginInvoke(New SetFolderNameTextCallback(AddressOf SetFoldersNameText), New Object() {Format(lTotalRecommendations, "###,####,###") & " e-mails"})

                        My.Settings.Save()

                    Else

                        ' Cancelled after ProcessAllFolders; show partial results if any
                        Collection_of_folders_to_exclude = Nothing
                        Collection_of_folders_to_exclude_is_empty = True
                        SetProcessPriorities("End Outlook Review")
                        MemoryManagement.FlushMemory()

                        ' hide and reset progress bar on cancel
                        Me.Dispatcher.BeginInvoke(New SetProgressBarVisableCallback(AddressOf SetProgressBarVisable),
                                             New Object() {Windows.Visibility.Hidden})
                        Me.Dispatcher.BeginInvoke(New SetProgressBarValueCallback(AddressOf SetProgressBarValue),
                                             New Object() {0.0R})

                        ' reset status-line email count on cancel
                        Me.Dispatcher.BeginInvoke(New SetFolderNameTextCallback(AddressOf SetFoldersNameText),
                                             New Object() {"0 e-mails"})

                    End If

                End If

                Me.Dispatcher.BeginInvoke(New SetToolTipCallback(AddressOf SetToolTip), New Object() {"Done"})

                SetProcessPriorities("End Review")

            Else

                Collection_of_folders_to_exclude = Nothing
                Collection_of_folders_to_exclude_is_empty = True

                SetProcessPriorities("End Outlook Review")

                MemoryManagement.FlushMemory()

                Me.Dispatcher.BeginInvoke(New SetFolderNameTextCallback(AddressOf SetFoldersNameText), New Object() {"0 e-mails"})

                Me.Dispatcher.BeginInvoke(New SetToolTipCallback(AddressOf SetToolTip), New Object() {"Done"})

                SetProcessPriorities("End Review")

            End If

            Me.Dispatcher.BeginInvoke(New FinalizeLoadCallback(AddressOf FinalizeLoad), New Object() {MSOutlookDrivenEvent})

CleanExit:

            If gCancelRefresh Then
                Me.Dispatcher.BeginInvoke(New ClearListView1Callback(AddressOf ClearListView1), New Object() {})
            End If


        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Refresh Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 ex.TargetSite.Name & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try

        MemoryManagement.FlushMemory()

#If DEBUG Then

        swOverall.Stop()
        Dim ts As TimeSpan = TimeSpan.FromMilliseconds(swOverall.ElapsedMilliseconds)
        Console.WriteLine($"Overall time to refresh: {ts.Hours} hours, {ts.Minutes} minutes, {ts.Seconds} seconds")
        Console.WriteLine("Total emails reviewed: " & lTotalEMailsToBeReviewed.ToString)
        Console.WriteLine("Emails / second: " & (lTotalEMailsToBeReviewed / (swOverall.ElapsedMilliseconds / 1000)).ToString("F2"))
        Console.WriteLine("")

#End If

    End Sub

#Region "Load Folder Table"

    Sub FindAllFolders()

        If gCancelRefresh Then Exit Sub

        gScanningFolders = True

        Try

            gFolderButtonsOnOptionsWindowEnabled = False

            gOptionsWindow?.SafelyEnableOptionsFolderButtons()

            ' Hide Mailbox column when there is only one mailbox
            UpdateMailboxColumnVisibility()

            gFolderTableIndex = 0
            lTotalEMails = 0
            lTotalEMailsToBeReviewed = 0

            If oNS.Folders IsNot Nothing Then
                For x As Integer = 1 To oNS.Folders.Count
                    If gCancelRefresh Then Exit For
                    AddFolder(oNS.Folders.Item(x))
                Next
            End If

            ' sw.Stop()
            'Console.WriteLine(sw.ElapsedMilliseconds.ToString) : sw.Stop()

            gFolderTableIndex -= 1
            gFolderTableCurrentSize = gFolderTableIndex

            ReDim Preserve gFolderTable(gFolderTableIndex)
            ReDim gFolderNamesTable(gFolderTableIndex)
            ReDim gFolderNamesTableTrimmed(gFolderTableIndex)

            For x As Integer = 0 To gFolderTable.Length - 1
                gFolderNamesTable(x) = gFolderTable(x).FolderPath
                gFolderNamesTableTrimmed(x) = gFolderNamesTable(x).TrimStart("\"c)
            Next

            ' Detect special folders across all mailboxes
            gDeletedFolderIndex = -1
            gInboxFolderIndices.Clear()
            gSentFolderIndices.Clear()
            gStoreDeleteFolders.Clear()

            ' First pass: locate Inbox/Sent and best delete folder (Deleted Items / Deleted / Trash) per store
            For x As Integer = 0 To gFolderTable.Length - 1

                Dim fInfo As FolderInfo = gFolderTable(x)
                If fInfo.DefaultItemType <> Microsoft.Office.Interop.Outlook.OlItemType.olMailItem Then
                    Continue For
                End If

                Dim nameUpper As String = System.IO.Path.GetFileName(fInfo.FolderPath).Trim().ToUpperInvariant()

                ' Track Inbox folders globally and per mailbox
                If nameUpper = "INBOX" Then
                    gInboxFolderIndices.Add(x)
                    Continue For
                End If

                ' Track Sent folders globally and per mailbox
                Select Case nameUpper
                    Case "SENT", "SENT ITEMS", "SENT MAIL"
                        gSentFolderIndices.Add(x)
                        Continue For
                End Select

                ' Figure out a suitable delete folder for this store:
                Dim isDeleted As Boolean
                Select Case nameUpper
                    Case "DELETED ITEMS", "DELETED", "TRASH"
                        isDeleted = True
                    Case Else
                        isDeleted = False
                End Select

                If Not isDeleted Then
                    Continue For
                End If

                Dim storeId As String = fInfo.StoreID
                If String.IsNullOrEmpty(storeId) Then
                    Continue For
                End If

                Dim existing As StoreDeleteFolderInfo = Nothing
                Dim hasExisting As Boolean = gStoreDeleteFolders.TryGetValue(storeId, existing)

                If Not hasExisting Then
                    ' First candidate for this store
                    gStoreDeleteFolders(storeId) = New StoreDeleteFolderInfo With {
                        .StoreId = storeId,
                        .FolderIndex = x
                    }
                Else
                    ' We may have a weaker candidate already; prefer Deleted Items > Deleted > Trash
                    Dim currentNameUpper As String = System.IO.Path.GetFileName(gFolderTable(existing.FolderIndex).FolderPath).Trim().ToUpperInvariant()

                    Dim currentScore As Integer = GetDeleteFolderPreferenceScore(currentNameUpper)
                    Dim newScore As Integer = GetDeleteFolderPreferenceScore(nameUpper)

                    If newScore > currentScore Then
                        existing.FolderIndex = x
                        gStoreDeleteFolders(storeId) = existing
                    End If
                End If

                ' Maintain backwards‑compatible global deleted index for legacy callers:
                ' pick the first 'Deleted Items' we see, then fall back to any previous.
                If (gDeletedFolderIndex = -1) AndAlso isDeleted Then
                    gDeletedFolderIndex = x
                End If

            Next

#If DEBUG Then
            ' Debug: ensure we found at least some inbox/sent folders
            Console.WriteLine("Inboxes: " & gInboxFolderIndices.Count & " Sent: " & gSentFolderIndices.Count)
#End If

            gFolderButtonsOnOptionsWindowEnabled = True
            gOptionsWindow?.SafelyEnableOptionsFolderButtons()

            Dim ToolTipMessage As String = ""
            Dim ProgressBarMaxValue As Double

            If gRefreshOtherFolders OrElse gRefreshInbox OrElse gRefreshSent Then

                If gRefreshOtherFolders Then

                    ToolTipMessage = "E-mails from all included folders are being reviewed"

                    'ProcessBarMaxValue = 
                    ' 10 times the TotalEMails To Be Reviewed for processing all info but the trailer +
                    ' 1 times the TotalEMails To Be Reviewed for processing the trailer + 
                    ' a time factor doing the recommendations
                    'ProgressBarMaxValue = (3 * lTotalEMailsToBeReviewed) + Int(lTotalEMailsToBeReviewed * (1 + My.Settings.RatioOfRecommendationToProcessingTime + 0.01))
                    ProgressBarMaxValue = lTotalEMailsToBeReviewed * (lProgressBarWeightingForEmailReviews + lProgressBarWeightingForFinalSteps)
                    ProgressBarMaxValue *= (1 + My.Settings.RatioOfRecommendationToProcessingTime + 0.01)

                Else

                    If gRefreshInbox AndAlso gRefreshSent Then
                        ToolTipMessage &= "Inbox and sent e-mails are being reviewed"

                    ElseIf gRefreshInbox Then
                        ToolTipMessage &= "Inbox e-mails are being reviewed"

                    Else
                        ToolTipMessage &= "Sent e-mails are being reviewed"

                    End If

                    ProgressBarMaxValue = lTotalEMailsToBeReviewed

                End If

            Else

                ToolTipMessage = "No e-mails are being reviewed based on the options chosen"
                lTotalEMailsToBeReviewed = 0
                gThresholdForReportingProgressOnTheProgressBar = 0

                ProgressBarMaxValue = 0

            End If

            gThresholdForReportingProgressOnTheProgressBar = lTotalEMailsToBeReviewed / 100 ' Math.Max(50, lTotalEMailsToBeReviewed / 100) ' every 1 percent

            Me.Dispatcher.BeginInvoke(New SetToolTipCallback(AddressOf SetToolTip), New Object() {ToolTipMessage})
            Me.Dispatcher.BeginInvoke(New SetProgressBarMaxValueCallback(AddressOf SetProgressBarMaxValue), New Object() {ProgressBarMaxValue})

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Find All Folders Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try

        gScanningFolders = False

    End Sub

    Private Function GetDeleteFolderPreferenceScore(ByVal nameUpper As String) As Integer

        ' Highest score wins
        If nameUpper = "DELETED ITEMS" Then
            Return 3
        End If
        If nameUpper = "DELETED" Then
            Return 2
        End If
        If nameUpper = "TRASH" Then
            Return 1
        End If

        Return 0

    End Function

    Private Sub AddFolder(ByRef StartFolder As Microsoft.Office.Interop.Outlook.MAPIFolder)

        If gCancelRefresh Then Exit Sub

        Dim defaultItemType As Microsoft.Office.Interop.Outlook.OlItemType
        Dim subFolders As Microsoft.Office.Interop.Outlook.Folders = Nothing

        Try

            defaultItemType = StartFolder.DefaultItemType

            If defaultItemType = Microsoft.Office.Interop.Outlook.OlItemType.olMailItem Then
                AddAnEntry(StartFolder)
            End If

            subFolders = StartFolder.Folders
            If subFolders Is Nothing Then Exit Sub

            Dim count As Integer = 0
            Try
                count = subFolders.Count
            Catch ex As System.Runtime.InteropServices.COMException
                Exit Sub
            Catch
                Exit Sub
            End Try

            For i As Integer = 1 To count

                If gCancelRefresh Then Exit For

                Dim oFolder As Microsoft.Office.Interop.Outlook.MAPIFolder = Nothing

                Try

                    Try
                        oFolder = subFolders.Item(i)
                    Catch ex As System.Runtime.InteropServices.COMException
                        Continue For
                    Catch
                        Continue For
                    End Try

                    If oFolder Is Nothing Then
                        Continue For
                    End If

                    Try
                        AddFolder(oFolder)
                    Catch ex As System.Runtime.InteropServices.COMException
                        ' Skip any sub-folder that errors
                    Catch
                        ' Ignore and continue with remaining sub-folders
                    End Try

                Finally

                    If oFolder IsNot Nothing Then
                        Try
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oFolder)
                        Catch
                        End Try
                        oFolder = Nothing
                    End If

                End Try

            Next

        Catch ex As System.Runtime.InteropServices.COMException
            ' Skip folders that cannot be inspected due to Outlook/MAPI errors
            Exit Sub
        Catch
            ' Any other error getting DefaultItemType – skip this folder
            Exit Sub
        Finally

            If subFolders IsNot Nothing Then
                Try
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(subFolders)
                Catch
                End Try
                subFolders = Nothing
            End If

        End Try

    End Sub

    Private Sub AddAnEntry(ByRef Folder As Microsoft.Office.Interop.Outlook.MAPIFolder)

        'Ensure the folder table Is initialized at least once
        If gFolderTable Is Nothing OrElse gFolderTable.Length = 0 Then
            gFolderTableCurrentSize = gFolderTableIncrement
            ReDim gFolderTable(gFolderTableCurrentSize)
        End If

        If gFolderTableIndex > gFolderTableCurrentSize - 1 Then
            gFolderTableCurrentSize += gFolderTableIncrement
            ReDim Preserve gFolderTable(gFolderTableCurrentSize)
        End If

        ' Store only folder identity data; do not keep COM objects across threads
        Dim info As FolderInfo
        info.EntryID = Folder.EntryID
        info.StoreID = Folder.StoreID
        info.FolderPath = Folder.FolderPath
        info.DefaultItemType = Folder.DefaultItemType

        gFolderTable(gFolderTableIndex) = info
        gFolderTableIndex += 1

        Dim CurrentFolderPath As String = info.FolderPath
        Dim Include As Boolean

        ' Determine if this folder is Inbox/Sent by its name
        Dim folderName As String = System.IO.Path.GetFileName(CurrentFolderPath).Trim()
        Dim isInboxFolder As Boolean = String.Equals(folderName, "Inbox", StringComparison.OrdinalIgnoreCase)
        Dim isSentFolder As Boolean = String.Equals(folderName, "Sent", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(folderName, "Sent Items", StringComparison.OrdinalIgnoreCase)

        If (gRefreshInbox AndAlso isInboxFolder) OrElse (gRefreshSent AndAlso isSentFolder) Then
            Include = True
        ElseIf gRefreshOtherFolders Then
            Include = Collection_of_folders_to_exclude_is_empty OrElse (Collection_of_folders_to_exclude.IndexOf(CurrentFolderPath) = -1)
        Else
            Include = False
        End If

        Dim folderItemCount As Integer = 0

        If Include Then
            Try
                folderItemCount = Folder.Items.Count
            Catch
                folderItemCount = 0
            End Try

            lTotalEMailsToBeReviewed += folderItemCount
        End If

        lTotalEMails += folderItemCount

        Dim msg As String
        If Include Then
            msg = "Including " & CurrentFolderPath.TrimStart("\"c)
        Else
            msg = "Excluding " & CurrentFolderPath.TrimStart("\"c)
        End If

        ' I tried throttling the UI updates below, but there was no noticeable performance gain
        Me.Dispatcher.BeginInvoke(
                New SetFolderNameTextCallback(AddressOf SetFoldersNameText),
                New Object() {msg})

    End Sub

#End Region

#Region "Load EMail Table"

    ' MD5-based fingerprint for trailers (restores original behavior)
    Private Function ComputeTrailerHash(ByVal value As String) As String

        If String.IsNullOrEmpty(value) Then
            Return ""
        End If

        Dim bytes As Byte() = System.Text.Encoding.ASCII.GetBytes(value)
        Dim hashBytes As Byte() = md5Obj.ComputeHash(bytes)

        ' Represent as 32-char hex string for stability / readability
        Dim sb As New System.Text.StringBuilder(hashBytes.Length * 2)
        For i As Integer = 0 To hashBytes.Length - 1
            sb.Append(hashBytes(i).ToString("X2", System.Globalization.CultureInfo.InvariantCulture))
        Next
        Return sb.ToString()
    End Function

    Private Function IsAMailboxBoxPath(ByVal s As String) As Boolean

        ' a mailbox path has exactly two backslashes - in the first two characters
        ' all other paths have more than two backslashes

        If String.IsNullOrEmpty(s) Then Return False

        For i As Integer = 2 To s.Length - 1
            If s.Chars(i) = "\"c Then
                Return False
            End If
        Next

        Return True

    End Function

    Private Sub ProcessAllFolders(ByVal MSOutlookDrivenEvent As Boolean, ByVal QuickRefresh As Boolean)

        Static LastFullyLoadedFolderTable() As FolderInfo
        Static LastFolderNamesTableTrimmed(0) As String

        Static LastEmailTable() As StructureOfEmailDetails
        Static LastEmailTableIndex As Integer = 0

        'Dim sw As New Stopwatch
        'sw.Start()

        '***************************************************************************
        'Step 1 initializations
        '***************************************************************************

        Dim iProgressBarValue As Double = 0

        lWhenSent = My.Settings.WhenSent

        gEmailTableIndex = 0

        ' Set the size of the gEmailTable based on the current estimate of emails to be reviewed
        ' Further resizing will be done later if needed

        If lTotalEMailsToBeReviewed <= 0 Then
            ReDim gEmailTable(0)
        Else
            ReDim gEmailTable(lTotalEMailsToBeReviewed)
        End If

        Dim strCollection = New System.Collections.Specialized.StringCollection
        strCollection = My.Settings.ExcludedScanFolders 'list of all folders to be excluded from scan

        With lBlankEMailDetailRecord

            .sSubject = ""
            .sTrailer = ""
            .sTo = ""
            .sFrom = ""
            .sDateAndTime = Now
            .sOutlookEntryID = ""
            .sTrailer = ""
            .sUnRead = System.Windows.FontWeights.Bold

        End With

        '***************************************************************************
        'Step 2 add all info except trailers 
        '***************************************************************************

        ' below we will add all other folders first, and inbox and sent items second
        ' this is done as the add of the other folders can be long running, while the inbox and sent items are quick
        ' in this way any emails sent or received during the long running other folder process will be picked up in the inbox and sent item processing that comes after

        Dim PopulateFoldersUsingPreviousData As Boolean = True

        Dim CurrentMailBoxName As String = ""

        If gRefreshOtherFolders AndAlso (Not MSOutlookDrivenEvent) AndAlso (Not QuickRefresh) Then

            For x As Integer = 0 To gFolderTableIndex

                If gCancelRefresh Then Exit For

                If IsAMailboxBoxPath(gFolderTable(x).FolderPath) Then
                    CurrentMailBoxName = GetMailboxNameFromFolderPath(gFolderTable(x).FolderPath, gFolderTable(x).StoreID)
                End If

                If (strCollection.IndexOf(gFolderNamesTable(x)) = -1) OrElse Collection_of_folders_to_exclude_is_empty Then

                    If gFolderTable(x).DefaultItemType = Microsoft.Office.Interop.Outlook.OlItemType.olMailItem Then
                        Dim folder As Microsoft.Office.Interop.Outlook.MAPIFolder = oNS.GetFolderFromID(gFolderTable(x).EntryID, gFolderTable(x).StoreID)
                        Try
                            ProcessAllMailItemsInAFolder(x, folder, CurrentMailBoxName, iProgressBarValue)
                        Finally
                            If folder IsNot Nothing Then
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(folder)
                                folder = Nothing
                            End If
                        End Try
                    End If

                End If

            Next

            ReDim LastFullyLoadedFolderTable(gFolderTable.Length - 1)
            Array.Copy(gFolderTable, LastFullyLoadedFolderTable, gFolderTable.Length)

            ReDim LastFolderNamesTableTrimmed(gFolderNamesTableTrimmed.Length - 1)
            Array.Copy(gFolderNamesTableTrimmed, LastFolderNamesTableTrimmed, gFolderNamesTableTrimmed.Length)

            ReDim LastEmailTable(gEmailTableIndex - 1)
            Array.Copy(gEmailTable, LastEmailTable, gEmailTableIndex)
            LastEmailTableIndex = gEmailTableIndex

            PopulateFoldersUsingPreviousData = False

        End If

        ' this keeps the last Fully Loaded Folder Table and the Email Table (with recommendations) in place when there is an MS Outlook driven event
        If PopulateFoldersUsingPreviousData Then

            If (LastFullyLoadedFolderTable IsNot Nothing) Then
                Array.Copy(LastFullyLoadedFolderTable, gFolderTable, LastFullyLoadedFolderTable.Length)
            End If

            If LastFolderNamesTableTrimmed IsNot Nothing Then
                Array.Copy(LastFolderNamesTableTrimmed, gFolderNamesTableTrimmed, LastFolderNamesTableTrimmed.Length)
            End If

            If LastEmailTable IsNot Nothing Then
                Array.Copy(LastEmailTable, gEmailTable, LastEmailTable.Length)
                gEmailTableIndex = LastEmailTableIndex
            End If

        End If

        If gRefreshInbox OrElse gRefreshSent Then

            For x As Int16 = 0 To gFolderTableIndex

                If gCancelRefresh Then Exit For

                If IsAMailboxBoxPath(gFolderTable(x).FolderPath) Then
                    CurrentMailBoxName = GetMailboxNameFromFolderPath(gFolderTable(x).FolderPath, gFolderTable(x).StoreID)
                End If

                If gRefreshSent AndAlso (gInboxFolderIndices.Count > 0) AndAlso gInboxFolderIndices.Contains(x) Then

                    If gFolderTable(x).DefaultItemType = Microsoft.Office.Interop.Outlook.OlItemType.olMailItem Then
                        Dim folder As Microsoft.Office.Interop.Outlook.MAPIFolder = oNS.GetFolderFromID(gFolderTable(x).EntryID, gFolderTable(x).StoreID)
                        Try
                            ProcessAllMailItemsInAFolder(x, folder, CurrentMailBoxName, iProgressBarValue)
                        Finally
                            If folder IsNot Nothing Then
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(folder)
                                folder = Nothing
                            End If
                        End Try
                    End If

                End If

                If gRefreshSent AndAlso (gSentFolderIndices.Count > 0) AndAlso gSentFolderIndices.Contains(x) Then

                    If gFolderTable(x).DefaultItemType = Microsoft.Office.Interop.Outlook.OlItemType.olMailItem Then
                        Dim folder As Microsoft.Office.Interop.Outlook.MAPIFolder = oNS.GetFolderFromID(gFolderTable(x).EntryID, gFolderTable(x).StoreID)
                        Try
                            ProcessAllMailItemsInAFolder(x, folder, CurrentMailBoxName, iProgressBarValue)
                        Finally
                            If folder IsNot Nothing Then
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(folder)
                                folder = Nothing
                            End If
                        End Try
                    End If

                End If

            Next

        End If

        If gCancelRefresh Then GoTo EarlyExit

        strCollection = Nothing

        ' resize email table to its actual needed size
        If gEmailTableIndex > 0 Then
            ReDim Preserve gEmailTable(gEmailTableIndex - 1)
        Else
            ReDim gEmailTable(0)
        End If

        '***************************************************************************
        'Step 2 sort email table so that subjects are grouped together
        '***************************************************************************

        Dim lEMailTableSorter As New EMailTableSorter With {
           .PrimarySortColumn = 1, ' subject
           .PrimaryOrder = SortOrder.Ascending,
           .SecondarySortColumn = 4, ' date and time
           .SecondaryOrder = SortOrder.Descending
           }

        Array.Sort(gEmailTable, lEMailTableSorter)
        lEMailTableSorter = Nothing

        '***************************************************************************
        'Step 3: trailer flags are used to identify email chains, they denote where
        '        an email's subject matches the subject of the next email; 
        '        add these trailers where needed
        '***************************************************************************

        Dim addTrailerFlag As Boolean

        For x = 0 To gEmailTable.Length - 1

            If gCancelRefresh Then Exit For

            Select Case x

                Case 0
                    If gEmailTable.Length > 1 Then
                        addTrailerFlag = (gEmailTable(x).sSubject = gEmailTable(x + 1).sSubject)
                    Else
                        addTrailerFlag = False
                    End If

                Case < gEmailTable.Length - 1
                    addTrailerFlag = (gEmailTable(x).sSubject = gEmailTable(x - 1).sSubject) OrElse
                                     (gEmailTable(x).sSubject = gEmailTable(x + 1).sSubject)

                Case Else
                    If gEmailTable.Length > 1 Then
                        addTrailerFlag = (gEmailTable(x).sSubject = gEmailTable(x - 1).sSubject)
                    Else
                        addTrailerFlag = False
                    End If

            End Select

            If addTrailerFlag Then

                ' the trailer is taken from the body of an email
                ' where the email is a standalone email, the trailer is taken from the top of the email, starting from the Subject line
                ' where the email is part of a chain of emails, the trailer starts from the Subject Line of the originating email in the email chain.
                ' the trailer field is then used to link emails together in chains

                If String.IsNullOrEmpty(gEmailTable(x).sBody) Then
                    gEmailTable(x).sTrailer = ""
                Else
                    gEmailTable(x).sTrailer = gEmailTable(x).sBody
                End If

                ' once the trailer is populated below the body is no longer needed and is cleared to save memory
                gEmailTable(x).sBody = Nothing

                gEmailTable(x).sTrailer = gEmailTable(x).sTrailer.Trim()

                If gEmailTable(x).sTrailer.Length = 0 Then

                    gEmailTable(x).sTrailer = Chr(255)

                Else

                    Dim trailer As String = gEmailTable(x).sTrailer

                    lLastIndex = trailer.LastIndexOf("Subject:", StringComparison.OrdinalIgnoreCase)
                    If lLastIndex > -1 Then
                        gEmailTable(x).sTrailer = trailer.Substring(lLastIndex + 8)
                    End If

                    ' Remove stuff so the email chains can be properly linked together
                    gEmailTable(x).sTrailer = gEmailTable(x).sTrailer _
                        .Replace(" ", "") _
                        .Replace(vbCr, "") _
                        .Replace(vbLf, "") _
                        .Replace(">", "") _
                        .Replace(vbTab, "")

                    Dim lHoldSubject As String =
                        gEmailTable(x).sSubject _
                            .Replace(" ", "") _
                            .Replace(vbCr, "") _
                            .Replace(vbLf, "") _
                            .Replace(">", "") _
                            .Replace(vbTab, "")

                    If lHoldSubject.Length > 0 Then
                        gEmailTable(x).sTrailer = gEmailTable(x).sTrailer.Replace(lHoldSubject, "")
                    End If

                    ' Only work with the first 256 characters of the trailer to avoid storing more information than is needed;
                    ' 256 characters was chosen as a reasonable compromise between uniqueness and performance/storage
                    If gEmailTable(x).sTrailer.Length > 256 Then
                        gEmailTable(x).sTrailer = gEmailTable(x).sTrailer.Remove(256)
                    End If

                    ' compute a hash of the trailer for storage in the email table
                    gEmailTable(x).sTrailer = ComputeTrailerHash(gEmailTable(x).sTrailer)

                End If

            End If

            lProgressBareRefreshingThresholdCounter += lProgressBarWeightingForFinalSteps
            If lProgressBareRefreshingThresholdCounter > gThresholdForReportingProgressOnTheProgressBar Then
                iProgressBarValue += lProgressBareRefreshingThresholdCounter
                lProgressBareRefreshingThresholdCounter = 0
                Me.Dispatcher.BeginInvoke(
                       New SetProgressBarValueCallback(AddressOf SetProgressBarValue),
                       New Object() {iProgressBarValue})
            End If

        Next

EarlyExit:

        'sw.Stop()
        'Console.WriteLine(sw.ElapsedMilliseconds)

    End Sub

    'moved here for performance gains

    Private lBlankEMailDetailRecord As StructureOfEmailDetails
    Private lWhenSent As Boolean
    Private lLastIndex As Integer

    Private FlipFlop As Boolean = True

    Private Sub ProcessAllMailItemsInAFolder(ByVal originalFolder As Int16,
                             ByVal folder As Microsoft.Office.Interop.Outlook.MAPIFolder,
                             ByVal mailboxName As String,
                             ByRef iProgressBarValue As Double)

        If gCancelRefresh Then Exit Sub

        Dim items As Microsoft.Office.Interop.Outlook.Items = Nothing

        Try
            items = folder.Items
        Catch
            Exit Sub
        End Try

        Try

            Try
                Dim sortField As String = If(lWhenSent, "[SentOn]", "[ReceivedTime]")
                items.Sort(sortField, True)
            Catch
            End Try

            Dim itemCount As Integer = items.Count ' set the number of items in the folder as a variable (to avoid having to access it repeatably in the line below)
            If itemCount = 0 Then Exit Sub

            ' Ensure there will be enough space in the email table when adding a new items
            If (gEmailTableIndex + itemCount) >= UBound(gEmailTable) Then
                ReDim Preserve gEmailTable(gEmailTableIndex + Math.Max(gEmailTableGrowth, itemCount))
            End If

            For Each item As Object In items   ' resting here change to for each item as object in items

                If gCancelRefresh Then Exit For

                Dim mail As Microsoft.Office.Interop.Outlook.MailItem = Nothing

                Try

                    mail = TryCast(item, Microsoft.Office.Interop.Outlook.MailItem)
                    If mail Is Nothing Then Continue For

                    Dim emailDetail As StructureOfEmailDetails = lBlankEMailDetailRecord

                    Dim friendlyFrom As String = mail.SenderEmailAddress

                    ' Resolve a friendly "From" address (gets around a quirk in Outlook / Exchange for messages coming from Exchange or certain connected accounts)
                    Try
                        If mail.Sender IsNot Nothing Then
                            ' If this is already an SMTP address, use it directly
                            If String.Equals(mail.SenderEmailType, "SMTP", StringComparison.OrdinalIgnoreCase) Then
                                ' all good
                            Else
                                ' Try to resolve to an Exchange user and use PrimarySmtpAddress
                                Dim exUser As Microsoft.Office.Interop.Outlook.ExchangeUser =
                                    TryCast(mail.Sender.GetExchangeUser(), Microsoft.Office.Interop.Outlook.ExchangeUser)

                                If exUser IsNot Nothing AndAlso Not String.IsNullOrEmpty(exUser.PrimarySmtpAddress) Then
                                    friendlyFrom = exUser.PrimarySmtpAddress

                                End If
                            End If

                        End If
                    Catch
                    End Try

                    With emailDetail

                        .sOriginalFolderReferenceNumber = originalFolder
                        .sOutlookEntryID = mail.EntryID
                        .sSubject = CleanUpSubjectLine(If(mail.Subject, String.Empty))
                        .sTo = If(mail.To, String.Empty)
                        .sFrom = If(friendlyFrom, String.Empty)
                        .sDateAndTime = If(lWhenSent, mail.SentOn, mail.ReceivedTime)
                        .sUnRead = If(mail.UnRead, System.Windows.FontWeights.Bold, System.Windows.FontWeights.Normal)
                        .sMailBoxName = mailboxName

                        ' Always capture body here; it may be trimmed/hashed later
                        .sBody = If(mail.Body, String.Empty)

                    End With

                    gEmailTable(gEmailTableIndex) = emailDetail
                    gEmailTableIndex += 1

                    lProgressBareRefreshingThresholdCounter += lProgressBarWeightingForEmailReviews
                    If lProgressBareRefreshingThresholdCounter > gThresholdForReportingProgressOnTheProgressBar Then
                        iProgressBarValue += lProgressBareRefreshingThresholdCounter
                        lProgressBareRefreshingThresholdCounter = 0
                        Me.Dispatcher.BeginInvoke(
                        New SetProgressBarValueCallback(AddressOf SetProgressBarValue),
                        New Object() {iProgressBarValue})
                    End If

                Catch
                Finally
                    If mail IsNot Nothing Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(mail)
                        mail = Nothing
                    End If
                    item = Nothing
                End Try
            Next

        Catch
        Finally
            If items IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(items)
            End If
        End Try

    End Sub


    Public Function CleanUpSubjectLine(ByVal subjectLine As String) As String

        subjectLine = subjectLine.Trim()

        Dim index = 0
        While index < subjectLine.Length - 2
            If (String.Compare(subjectLine, index, "RE:", 0, 3, StringComparison.OrdinalIgnoreCase) = 0 OrElse
            String.Compare(subjectLine, index, "FW:", 0, 3, StringComparison.OrdinalIgnoreCase) = 0) Then
                index += 3
                While index < subjectLine.Length AndAlso Char.IsWhiteSpace(subjectLine(index))
                    index += 1
                End While
            Else
                Exit While
            End If
        End While

        Dim result = If(index > 0, subjectLine.Substring(index), subjectLine)

        If String.IsNullOrWhiteSpace(result) Then
            Return " "
        Else
            Return result
        End If

    End Function

#End Region

#Region "Establish and set rankings"

    Private Sub EstablishRecommendations()

        Try

            'A second sort of the email table is required to subjects in order with their trailers

            Dim lEMailTableSorter As New EMailTableSorter With {
            .PrimarySortColumn = 1, ' subject
            .PrimaryOrder = SortOrder.Ascending,
            .SecondarySortColumn = 4, ' date and time 
            .SecondaryOrder = SortOrder.Descending
            }
            Array.Sort(gEmailTable, lEMailTableSorter)
            lEMailTableSorter = Nothing

            EstablishRatings_NumberOfUniqueEmailsInAFolder()

            EstablishRatings_Scoring()

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Establish Recommendations Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 ex.TargetSite.Name & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try

    End Sub

    Private Sub EstablishRatings_NumberOfUniqueEmailsInAFolder()

        'Set up for rating number of e-mails related to the same chain within a folder

        Try
            UniqueSubjectsMap.Clear()

            For x As Int32 = 0 To gEmailTable.Length - 1
                Dim subjectAndTrailer As String = gEmailTable(x).sSubject & gEmailTable(x).sTrailer
                Dim folderRef As Integer = gEmailTable(x).sOriginalFolderReferenceNumber

                Dim folderCounts As Dictionary(Of Integer, Int16) = Nothing
                If Not UniqueSubjectsMap.TryGetValue(subjectAndTrailer, folderCounts) Then
                    folderCounts = New Dictionary(Of Integer, Int16)()
                    UniqueSubjectsMap(subjectAndTrailer) = folderCounts
                End If

                Dim count As Int16
                If folderCounts.TryGetValue(folderRef, count) Then
                    folderCounts(folderRef) = CType(count + 1, Int16)
                Else
                    folderCounts(folderRef) = 1
                End If
            Next

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Establish Ratings Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 ex.TargetSite.Name & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try

    End Sub

    Private Sub EstablishRatings_Scoring()

        'For each unique email chain, rate the best folder to put it in
        '   1 point to each folder for each email in it that belongs to the same unique email chain

        Try

            Dim CurrentSubjectAndTrailer As String = "|*| something unique |*|" & Chr(255)
            Dim PrevSubjectAndTrailer As String = ""

            Dim FinalScoringTable(gFolderTable.Length - 1) As Integer

            For i As Integer = 0 To gEmailTable.Length - 1

                PrevSubjectAndTrailer = CurrentSubjectAndTrailer
                CurrentSubjectAndTrailer = gEmailTable(i).sSubject & gEmailTable(i).sTrailer

                If CurrentSubjectAndTrailer <> PrevSubjectAndTrailer Then

                    Dim folderCounts As Dictionary(Of Integer, Int16) = Nothing
                    UniqueSubjectsMap.TryGetValue(CurrentSubjectAndTrailer, folderCounts)

                    Array.Clear(FinalScoringTable, 0, FinalScoringTable.Length)

                    If folderCounts IsNot Nothing Then
                        For Each kvp As KeyValuePair(Of Integer, Int16) In folderCounts
                            If kvp.Key >= 0 AndAlso kvp.Key < FinalScoringTable.Length Then
                                FinalScoringTable(kvp.Key) = kvp.Value
                            End If
                        Next
                    End If

                    ' Don't recommend the original folder
                    'FinalScoringTable(gEmailTable(x).sOriginalFolderReferenceNumber) = 0

                    ' Don't recommend any inbox or sent items
                    For Each idx As Integer In gInboxFolderIndices
                        If idx >= 0 AndAlso idx < FinalScoringTable.Length Then
                            FinalScoringTable(idx) = 0
                        End If
                    Next
                    For Each idx As Integer In gSentFolderIndices
                        If idx >= 0 AndAlso idx < FinalScoringTable.Length Then
                            FinalScoringTable(idx) = 0
                        End If
                    Next

                    FindTheFolderWithTheGreatestScore(gEmailTable(i).sRecommendedFolder1ReferenceNumber, FinalScoringTable)
                    FindTheFolderWithTheGreatestScore(gEmailTable(i).sRecommendedFolder2ReferenceNumber, FinalScoringTable)
                    FindTheFolderWithTheGreatestScore(gEmailTable(i).sRecommendedFolder3ReferenceNumber, FinalScoringTable)
                    gEmailTable(i).sRecommendedFolderFinalReferenceNumber = gEmailTable(i).sRecommendedFolder1ReferenceNumber

                Else

                    gEmailTable(i).sRecommendedFolder1ReferenceNumber = gEmailTable(i - 1).sRecommendedFolder1ReferenceNumber
                    gEmailTable(i).sRecommendedFolder2ReferenceNumber = gEmailTable(i - 1).sRecommendedFolder2ReferenceNumber
                    gEmailTable(i).sRecommendedFolder3ReferenceNumber = gEmailTable(i - 1).sRecommendedFolder3ReferenceNumber
                    gEmailTable(i).sRecommendedFolderFinalReferenceNumber = gEmailTable(i - 1).sRecommendedFolderFinalReferenceNumber

                End If

            Next

            UniqueSubjectsMap.Clear()

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Establish Ratings Scoring Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 ex.TargetSite.Name & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try


    End Sub

    Private Sub FindTheFolderWithTheGreatestScore(ByRef ReferenceNumber As Integer, ByRef FinalScoringTable() As Integer)

        'The following code works with dot net version 3.5 only, is replaced by the code block below it for use with dot net version 3.0
        'Find the entryId with the greatest final score
        'Dim Max As Integer = FinalScoringTable.Take(FinalScoringTable.Length).Max()
        'Dim MaxIndex As Integer = Array.IndexOf(FinalScoringTable, Max)

        Dim max As Integer = 0
        Dim MaxIndex As Integer = 0
        For x As Integer = 0 To FinalScoringTable.Length - 1
            If FinalScoringTable(x) > max Then
                max = FinalScoringTable(x)
                MaxIndex = x
            End If
        Next

        'Assign it as the winner, reset its score to zero so it can't win again
        If max > 0 Then
            ReferenceNumber = MaxIndex
            FinalScoringTable(MaxIndex) = 0
        Else
            ReferenceNumber = -1
        End If

    End Sub

#End Region

    Private Sub UpdateListView()

        Try
            Dim operation1 = Me.Dispatcher.BeginInvoke(Sub() StorePendingSelection(SelectionRestoreReason.Refresh))
            operation1.Wait()
        Catch
        End Try

        Try

            '' for debugging print the contents of the email table to the console for debugging
            'For x As Integer = 0 To gEmailTableIndex - 1
            '    Debug.WriteLine(x.ToString & " " & gEmailTable(x).sSubject)
            'Next

            Dim operation2 = Me.Dispatcher.BeginInvoke(New ClearListView1Callback(AddressOf ClearListView1), New Object() {})
            operation2.Wait()

            If gEmailTableIndex = 0 Then Exit Try

            ReDim gFinalRecommendationTable(gEmailTableIndex)

            Dim lLineNumber As Integer = 0
            Dim lNextIndex As Integer = 0

            Dim lCurrentSubjectPlusTrailer As String = ""
            Dim lPreviousSubjectPlusTrailer As String = ""

            Dim lFirstSubjectPlusTrailer As String = ""
            Dim lNextSubjectPlusTrailer As String = ""

            Dim lRecommendedIndexForAllEntriesInChainFinal As Integer
            Dim lRecommendedIndexForAllEntriesInChain1 As Integer
            Dim lRecommendedIndexForAllEntriesInChain2 As Integer
            Dim lRecommendedIndexForAllEntriesInChain3 As Integer

            Dim lFlagThisEmailChain As Boolean

            For x As Integer = 0 To gEmailTableIndex - 1

                lRecommendedIndexForAllEntriesInChainFinal = -1
                lRecommendedIndexForAllEntriesInChain1 = -1
                lRecommendedIndexForAllEntriesInChain2 = -1
                lRecommendedIndexForAllEntriesInChain3 = -1

                lFlagThisEmailChain = False

                lNextIndex = x
                lFirstSubjectPlusTrailer = gEmailTable(x).sSubject & gEmailTable(x).sTrailer
                lNextSubjectPlusTrailer = lFirstSubjectPlusTrailer

                ' for each email chain
                ' flag the chain to be reported if it contains an inbox item, sent item, or and email store anyplace other than the recommended folder
                While (lFirstSubjectPlusTrailer = lNextSubjectPlusTrailer) And (lNextIndex <= (gEmailTableIndex - 1))

                    If gInboxFolderIndices.Contains(gEmailTable(lNextIndex).sOriginalFolderReferenceNumber) Then

                        If gRefreshInbox Then lFlagThisEmailChain = True

                    ElseIf gSentFolderIndices.Contains(gEmailTable(lNextIndex).sOriginalFolderReferenceNumber) Then

                        If gRefreshSent Then lFlagThisEmailChain = True

                    Else

                        If gEmailTable(lNextIndex).sRecommendedFolderFinalReferenceNumber > -1 Then

                            If gEmailTable(lNextIndex).sOriginalFolderReferenceNumber <> gEmailTable(lNextIndex).sRecommendedFolderFinalReferenceNumber Then

                                lRecommendedIndexForAllEntriesInChainFinal = gEmailTable(lNextIndex).sRecommendedFolderFinalReferenceNumber
                                lRecommendedIndexForAllEntriesInChain1 = gEmailTable(lNextIndex).sRecommendedFolder1ReferenceNumber
                                lRecommendedIndexForAllEntriesInChain2 = gEmailTable(lNextIndex).sRecommendedFolder2ReferenceNumber
                                lRecommendedIndexForAllEntriesInChain3 = gEmailTable(lNextIndex).sRecommendedFolder3ReferenceNumber
                                lFlagThisEmailChain = True
                                Exit While

                            End If
                        End If

                    End If

                    lNextIndex += 1
                    If (lNextIndex <= (gEmailTableIndex - 1)) Then
                        lNextSubjectPlusTrailer = gEmailTable(lNextIndex).sSubject & gEmailTable(lNextIndex).sTrailer
                    End If

                End While

                'ensure if an e-mail chain is flagged the a recommendation is made if at all possible
                'the following covers the case where there are inbox or sent items and all filed emails are in the same folder
                If lFlagThisEmailChain Then

                    If lRecommendedIndexForAllEntriesInChainFinal = -1 Then

                        lNextIndex = x
                        lFirstSubjectPlusTrailer = gEmailTable(x).sSubject & gEmailTable(x).sTrailer
                        lNextSubjectPlusTrailer = lFirstSubjectPlusTrailer

                        While (lFirstSubjectPlusTrailer = lNextSubjectPlusTrailer) And (lNextIndex <= (gEmailTableIndex - 1))

                            If Not gInboxFolderIndices.Contains(gEmailTable(lNextIndex).sOriginalFolderReferenceNumber) Then

                                If Not gSentFolderIndices.Contains(gEmailTable(lNextIndex).sOriginalFolderReferenceNumber) Then

                                    lRecommendedIndexForAllEntriesInChainFinal = gEmailTable(lNextIndex).sRecommendedFolderFinalReferenceNumber
                                    lRecommendedIndexForAllEntriesInChain1 = gEmailTable(lNextIndex).sRecommendedFolder1ReferenceNumber
                                    lRecommendedIndexForAllEntriesInChain2 = gEmailTable(lNextIndex).sRecommendedFolder2ReferenceNumber
                                    lRecommendedIndexForAllEntriesInChain3 = gEmailTable(lNextIndex).sRecommendedFolder3ReferenceNumber
                                    lFlagThisEmailChain = True
                                    Exit While

                                End If

                            End If

                            lNextIndex += 1

                            If (lNextIndex <= (gEmailTableIndex - 1)) Then
                                lNextSubjectPlusTrailer = gEmailTable(lNextIndex).sSubject & gEmailTable(lNextIndex).sTrailer
                            End If


                        End While

                    End If

                End If


                If lFlagThisEmailChain Then

                    Dim lStartingSubjectPlusTrailer As String = gEmailTable(x).sSubject & gEmailTable(x).sTrailer
                    Dim lChainEntry As Integer = x

                    While lStartingSubjectPlusTrailer = gEmailTable(lChainEntry).sSubject & gEmailTable(lChainEntry).sTrailer

                        With gEmailTable(lChainEntry)

                            gFinalRecommendationTable(lLineNumber) = New ListViewRowClass
                            gFinalRecommendationTable(lLineNumber).Index = lChainEntry
                            gFinalRecommendationTable(lLineNumber).MailBoxName = .sMailBoxName
                            gFinalRecommendationTable(lLineNumber).Subject = .sSubject
                            gFinalRecommendationTable(lLineNumber).Trailer = .sTrailer
                            gFinalRecommendationTable(lLineNumber).From = .sFrom
                            gFinalRecommendationTable(lLineNumber).xTo = .sTo
                            gFinalRecommendationTable(lLineNumber).DateTime = .sDateAndTime
                            gFinalRecommendationTable(lLineNumber).OriginalFolder = .sOriginalFolderReferenceNumber
                            gFinalRecommendationTable(lLineNumber).RecommendedFolder1 = lRecommendedIndexForAllEntriesInChain1
                            gFinalRecommendationTable(lLineNumber).RecommendedFolder2 = lRecommendedIndexForAllEntriesInChain2
                            gFinalRecommendationTable(lLineNumber).RecommendedFolder3 = lRecommendedIndexForAllEntriesInChain3
                            gFinalRecommendationTable(lLineNumber).RecommendedFolderFinal = lRecommendedIndexForAllEntriesInChainFinal
                            gFinalRecommendationTable(lLineNumber).OutlookEntryID = .sOutlookEntryID
                            gFinalRecommendationTable(lLineNumber).UnRead = .sUnRead
                            lLineNumber += 1

                        End With

                        lChainEntry += 1
                        If lChainEntry > gEmailTableIndex - 1 Then
                            Exit While
                        End If

                    End While

                    x = lChainEntry - 1

                End If

            Next

            ReDim Preserve gFinalRecommendationTable(lLineNumber - 1)

            ApplyCurrentSortOrderToFinalTable()

            Dim operation5 = Me.Dispatcher.BeginInvoke(New SetListViewItemCallback(AddressOf SetListViewItem), New Object() {gFinalRecommendationTable})
            operation5.Wait()

            lTotalRecommendations = lLineNumber

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Update List View Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 ex.TargetSite.Name & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try

        Try
            ' Defer restore during a full refresh; ApplyFilter will restore once the final list is built, avoiding double invocation.
            If Not gIsRefreshing Then
                Dim operation1 = Me.Dispatcher.BeginInvoke(New RestoreSelectionCallback(AddressOf RestorePendingSelection), New Object() {})
                operation1.Wait()
            End If
        Catch
        End Try

    End Sub

    Private Sub MainWindow_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseLeftButtonDown

        Try
            gWhoIsInControl = WhoIsInControlType.Main
            DragMove()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub MainWindow_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Me.KeyUp, Menu1.KeyUp
        ProcessKeyUp(e)
    End Sub

    Private Sub MainWindow_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Me.KeyDown, Menu1.KeyDown
        ProcessKeyDown(e)
    End Sub

    Private Sub ListView1_ContextMenuOpening(ByVal sender As Object,
                                         ByVal e As ContextMenuEventArgs) _
                                         Handles ListView1.ContextMenuOpening

        If ListView1.SelectedItem Is Nothing Then
            If Me.MenuContextToggleRead IsNot Nothing Then
                Me.MenuContextToggleRead.Visibility = Windows.Visibility.Collapsed
            End If
        Else
            SafelyUpdateReadToggleContextMenu()
        End If

    End Sub

    Private Sub ListView1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ListView1.MouseDoubleClick
        'ignore double click if it happened within the bounds of the scroll bar

        If ListView1.ActualWidth - e.GetPosition(Me.ListView1).X > 15 Then

            ' open the e‑mail associated with the clicked entryId
            ' (matches what the "Open" menu/command does, but without asking for confirmation)
            If ConfirmActionMessage("Open") Then
                OpenAnEmail()
            End If

        End If

    End Sub

    Private Sub MainWindow_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseDown
        MenuKeyStrokeOverRide = False
    End Sub

    Private Sub ListView1_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles ListView1.MouseEnter

        If Me.Cursor IsNot Cursors.Wait Then
            Me.Cursor = Cursors.Arrow
        End If

    End Sub

    Private Sub ListView1_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles ListView1.MouseLeave

        If Me.Cursor IsNot Cursors.Wait Then
            Me.Cursor = Cursors.Arrow
        End If

    End Sub

    Private Sub ListView1_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ListView1.SelectionChanged

        If gSuppressUpdatesToDetailBox Then Exit Sub

        'Select all emails in chain if click happened on a chain indicator
        Dim MouseY As Integer = System.Windows.Forms.Control.MousePosition.Y
        If MouseY >= Me.Top AndAlso MouseY <= Me.Top + Me.ActualHeight Then
            Dim MouseX As Integer = System.Windows.Forms.Control.MousePosition.X
            Dim LeftBound As Integer = MouseX - Me.Left - 15
            Dim RightBound As Integer = Me.Left + Me.ListView1.ActualWidth - MouseX

            If (LeftBound < 20) OrElse (RightBound < 18) Then
                SelectAllMembersOfAnEmailChain()
            End If
        End If

        ' When AutoChainSelect is enabled, any selection change should expand to the full chain
        If gAutoChainSelect Then
            RemoveHandler ListView1.SelectionChanged, AddressOf ListView1_SelectionChanged
            Try
                SelectAllMembersOfAnEmailChain()
            Finally
                AddHandler ListView1.SelectionChanged, AddressOf ListView1_SelectionChanged
            End Try
        End If

        UpdateDetails()

    End Sub

    Private Function BuildChainKey(ByVal row As ListViewRowClass) As String

        If row Is Nothing Then Return ""
        Return (If(row.Subject, "") & If(row.Trailer, ""))

    End Function
    Private Function CaptureSelectionSnapshot() As SelectionSnapshot

        Dim snap As New SelectionSnapshot With {
        .Entries = New List(Of SelectionEntry),
        .FirstIndex = 0
        }

        For Each obj In ListView1.SelectedItems
            Dim row = TryCast(obj, ListViewRowClass)
            If row Is Nothing Then Continue For

            Dim entry As New SelectionEntry With {
                .OutlookEntryId = row.OutlookEntryID,
                .ChainKey = BuildChainKey(row),
                .Index = row.Index
            }
            snap.Entries.Add(entry)
        Next

        If snap.Entries.Count > 0 Then
            snap.HasSelection = True
            snap.FirstIndex = snap.Entries(0).Index
            For Each entry In snap.Entries
                If entry.Index < snap.FirstIndex Then
                    snap.FirstIndex = entry.Index
                End If
            Next
        Else
            snap.HasSelection = False
        End If

        Return snap

    End Function

    Private Sub StorePendingSelection(ByVal reason As SelectionRestoreReason)

        gPendingSelectionSnapshot = CaptureSelectionSnapshot()
        gPendingSelectionReason = reason
        gPendingSelectionFallbackToFirst = (gPendingSelectionSnapshot Is Nothing OrElse Not gPendingSelectionSnapshot.HasSelection)
        gPendingSelectionApplied = False

    End Sub

    Delegate Sub RestoreSelectionCallback()
    Private Sub RestorePendingSelection()

        ' Prevent redundant re-entry only when a prior restore already produced a selection; allow a second pass if nothing was selected yet.
        If gPendingSelectionApplied AndAlso ListView1 IsNot Nothing AndAlso ListView1.SelectedItems.Count > 0 Then Return

        RestoreSelection(gPendingSelectionSnapshot, gPendingSelectionReason, gPendingSelectionFallbackToFirst)
        gPendingSelectionSnapshot = Nothing

    End Sub
    Private Sub RestoreSelection(ByVal snapshot As SelectionSnapshot, ByVal reason As SelectionRestoreReason, ByVal fallbackToFirst As Boolean)
        'Restore selection state after listview updates, honoring deletion/sort reasons and chaining rules.

        If ListView1 Is Nothing Then Return

        'Temporarily detach selection handler to avoid reentrancy while rebuilding selection.
        RemoveHandler ListView1.SelectionChanged, AddressOf ListView1_SelectionChanged

        Try

            'Reset selection and bail out quickly if the list is empty.
            ListView1.SelectedItems.Clear()

            If ListView1.Items.Count = 0 Then
                gCurrentlySelectedListViewItemIndex = 0
                gPendingSelectionApplied = True
                BlankOutDetails()
                UpdateMainMessageLine()
                Return
            End If

            'When items were removed, pick the closest surviving entry and reselect chain members if enabled.
            If reason = SelectionRestoreReason.UserDelete OrElse reason = SelectionRestoreReason.OutlookDelete Then

                Dim targetIndex As Integer = -1
                If snapshot IsNot Nothing AndAlso snapshot.HasSelection Then
                    Dim minIndex As Integer = Integer.MaxValue
                    For Each entry In snapshot.Entries
                        If entry.Index < minIndex Then minIndex = entry.Index
                    Next
                    If minIndex <> Integer.MaxValue Then
                        targetIndex = minIndex
                    End If
                End If

                If targetIndex < 0 Then targetIndex = 0
                If targetIndex > ListView1.Items.Count - 1 Then targetIndex = ListView1.Items.Count - 1

                If targetIndex >= 0 AndAlso targetIndex < ListView1.Items.Count Then
                    If gAutoChainSelect Then
                        Dim anchorRow = TryCast(ListView1.Items(targetIndex), ListViewRowClass)
                        Dim key As String = BuildChainKey(anchorRow)
                        If key IsNot Nothing Then
                            For i As Integer = 0 To ListView1.Items.Count - 1
                                Dim row = TryCast(ListView1.Items(i), ListViewRowClass)
                                If row Is Nothing Then Continue For
                                If String.Equals(BuildChainKey(row), key, StringComparison.OrdinalIgnoreCase) Then
                                    ListView1.SelectedItems.Add(ListView1.Items(i))
                                End If
                            Next
                        Else
                            ListView1.SelectedItems.Add(ListView1.Items(targetIndex))
                        End If
                        ListView1.SelectedIndex = targetIndex
                    Else
                        ListView1.SelectedIndex = targetIndex
                    End If
                    gCurrentlySelectedListViewItemIndex = targetIndex
                Else
                    gCurrentlySelectedListViewItemIndex = 0
                    ListView1.SelectedIndex = -1
                End If

                If ListView1.SelectedIndex >= 0 Then
                    gPendingSelectionApplied = True
                    ListView1.UpdateLayout()
                    Dim selectedItem = ListView1.SelectedItem
                    ListView1.ScrollIntoView(selectedItem)
                    Dispatcher.BeginInvoke(Sub()
                                               Dim selectedContainer = TryCast(ListView1.ItemContainerGenerator.ContainerFromIndex(ListView1.SelectedIndex), System.Windows.Controls.ListViewItem)
                                               If selectedContainer IsNot Nothing Then
                                                   selectedContainer.Focus()
                                               Else
                                                   ListView1.Focus()
                                               End If
                                           End Sub, System.Windows.Threading.DispatcherPriority.Background)
                End If

                UpdateDetails()
                Return

            End If

            'If nothing to restore, optionally default to the first item and chain mates.
            If snapshot Is Nothing OrElse Not snapshot.HasSelection Then

                If fallbackToFirst Then
                    If ListView1.Items.Count > 0 Then
                        ListView1.SelectedIndex = 0
                        gCurrentlySelectedListViewItemIndex = ListView1.SelectedIndex
                        If gAutoChainSelect Then
                            Dim chainRow = TryCast(ListView1.SelectedItem, ListViewRowClass)
                            Dim chainKey = BuildChainKey(chainRow)
                            If chainKey IsNot Nothing Then
                                Dim extraIndices As New List(Of Integer)
                                For i As Integer = 0 To ListView1.Items.Count - 1
                                    Dim row = TryCast(ListView1.Items(i), ListViewRowClass)
                                    If row Is Nothing Then Continue For
                                    If String.Equals(BuildChainKey(row), chainKey, StringComparison.OrdinalIgnoreCase) Then
                                        extraIndices.Add(i)
                                    End If
                                Next

                                For Each idx In extraIndices
                                    If Not ListView1.SelectedItems.Contains(ListView1.Items(idx)) Then
                                        ListView1.SelectedItems.Add(ListView1.Items(idx))
                                    End If
                                Next

                            End If
                        End If
                    End If
                End If

                If ListView1.SelectedIndex >= 0 Then
                    gPendingSelectionApplied = True
                End If

                UpdateDetails()
                Return

            End If

            'Map original snapshot indices and Outlook IDs to current list indices for matching.
            Dim sourceIndexToListViewIndex As New Dictionary(Of Integer, Integer)
            Dim idToIndex As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

            For i As Integer = 0 To ListView1.Items.Count - 1
                Dim row = TryCast(ListView1.Items(i), ListViewRowClass)
                If row Is Nothing Then Continue For

                sourceIndexToListViewIndex.Add(row.Index, i)

                If Not String.IsNullOrEmpty(row.OutlookEntryID) Then
                    If Not idToIndex.ContainsKey(row.OutlookEntryID) Then
                        idToIndex.Add(row.OutlookEntryID, i)
                    End If
                End If
            Next

            'Track chain keys from the snapshot to reapply chain selections when appropriate.
            Dim selectedChainKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim selectedIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each entry In snapshot.Entries
                If Not String.IsNullOrEmpty(entry.ChainKey) Then
                    selectedChainKeys.Add(entry.ChainKey)
                End If
                If Not String.IsNullOrEmpty(entry.OutlookEntryId) Then
                    selectedIds.Add(entry.OutlookEntryId)
                End If
            Next

            'Resolve which current rows should be selected and where the anchor should be.
            Dim targetIndices As New HashSet(Of Integer)
            Dim anchorIndex As Integer = -1
            Dim maxSnapshotIndex As Integer = -1

            For Each entry In snapshot.Entries
                If entry.Index > maxSnapshotIndex Then maxSnapshotIndex = entry.Index
                Dim idx As Integer

                ' Prefer matching by Outlook ID; only fall back to index when no ID is available for the entry.
                If Not String.IsNullOrEmpty(entry.OutlookEntryId) AndAlso idToIndex.TryGetValue(entry.OutlookEntryId, idx) Then
                    targetIndices.Add(idx)
                ElseIf String.IsNullOrEmpty(entry.OutlookEntryId) AndAlso sourceIndexToListViewIndex.TryGetValue(entry.Index, idx) Then
                    targetIndices.Add(idx)
                End If
            Next

            'Ensure all surviving rows with matching Outlook IDs from the snapshot are reselected, even if indices shifted.
            If selectedIds.Count > 0 Then
                For i As Integer = 0 To ListView1.Items.Count - 1
                    Dim row = TryCast(ListView1.Items(i), ListViewRowClass)
                    If row Is Nothing Then Continue For
                    If Not String.IsNullOrEmpty(row.OutlookEntryID) AndAlso selectedIds.Contains(row.OutlookEntryID) Then
                        targetIndices.Add(i)
                    End If
                Next
            End If


            'Expand selection to other rows in the same chain when appropriate and chaining is enabled.
            If (reason <> SelectionRestoreReason.UserDelete AndAlso reason <> SelectionRestoreReason.OutlookDelete AndAlso reason <> SelectionRestoreReason.Refresh) AndAlso gAutoChainSelect AndAlso selectedChainKeys.Count > 0 Then
                For i As Integer = 0 To ListView1.Items.Count - 1
                    Dim row = TryCast(ListView1.Items(i), ListViewRowClass)
                    If row Is Nothing Then Continue For

                    If selectedChainKeys.Contains(BuildChainKey(row)) Then
                        targetIndices.Add(i)
                    End If
                Next
            End If

            'During sort restores, keep only contiguous chain selections to mirror visual grouping.
            If reason = SelectionRestoreReason.Sort AndAlso selectedChainKeys.Count > 0 AndAlso targetIndices.Count > 1 Then
                For Each key In selectedChainKeys
                    Dim indices As New List(Of Integer)
                    For Each idx In targetIndices
                        Dim row = TryCast(ListView1.Items(idx), ListViewRowClass)
                        If row Is Nothing Then Continue For
                        If String.Equals(BuildChainKey(row), key, StringComparison.OrdinalIgnoreCase) Then
                            indices.Add(idx)
                        End If
                    Next

                    indices.Sort()

                    If indices.Count > 1 Then
                        Dim contiguous As Boolean = True
                        For i As Integer = 1 To indices.Count - 1
                            If indices(i) <> indices(i - 1) + 1 Then
                                contiguous = False
                                Exit For
                            End If
                        Next

                        If Not contiguous Then
                            For i As Integer = 1 To indices.Count - 1
                                targetIndices.Remove(indices(i))
                            Next
                        End If
                    End If
                Next
            End If

            'If no exact matches remain, derive a reasonable fallback index and reapply chaining if available.
            If targetIndices.Count = 0 AndAlso snapshot.HasSelection Then
                Dim fallbackIndex As Integer
                If reason = SelectionRestoreReason.UserDelete OrElse reason = SelectionRestoreReason.OutlookDelete Then
                    fallbackIndex = maxSnapshotIndex
                    If fallbackIndex < 0 Then fallbackIndex = snapshot.FirstIndex
                Else
                    fallbackIndex = snapshot.FirstIndex - 1
                End If
                If fallbackIndex < 0 Then fallbackIndex = 0
                If fallbackIndex > ListView1.Items.Count - 1 Then fallbackIndex = ListView1.Items.Count - 1

                If fallbackIndex >= 0 AndAlso fallbackIndex < ListView1.Items.Count Then
                    Dim row = TryCast(ListView1.Items(fallbackIndex), ListViewRowClass)
                    If gAutoChainSelect AndAlso row IsNot Nothing Then
                        Dim key As String = BuildChainKey(row)
                        For i As Integer = 0 To ListView1.Items.Count - 1
                            Dim r = TryCast(ListView1.Items(i), ListViewRowClass)
                            If r Is Nothing Then Continue For
                            If String.Equals(BuildChainKey(r), key, StringComparison.OrdinalIgnoreCase) Then
                                targetIndices.Add(i)
                            End If
                        Next
                    Else
                        targetIndices.Add(fallbackIndex)
                    End If
                    anchorIndex = fallbackIndex
                End If
            End If

            ' When refreshing, if every previously selected row disappeared, default to the first remaining row.
            If targetIndices.Count = 0 AndAlso reason = SelectionRestoreReason.Refresh AndAlso snapshot IsNot Nothing AndAlso snapshot.HasSelection AndAlso ListView1.Items.Count > 0 Then
                fallbackToFirst = True
            End If

            'As a last resort, optionally select the first item.
            If targetIndices.Count = 0 AndAlso fallbackToFirst Then
                If ListView1.Items.Count > 0 Then
                    targetIndices.Add(0)
                End If
            End If

            'Apply the resolved selection, ensuring a stable anchor index.
            If targetIndices.Count > 0 Then
                Dim ordered As New List(Of Integer)(targetIndices)
                ordered.Sort()

                ' Guard against stale indices from the snapshot; skip anything outside the current bounds before applying.
                ordered = ordered.Where(Function(i) i >= 0 AndAlso i < ListView1.Items.Count).ToList()

                If ordered.Count > 0 Then
                    For Each idx In ordered
                        ListView1.SelectedItems.Add(ListView1.Items(idx))
                    Next
                    If anchorIndex < 0 AndAlso ordered.Count > 0 Then anchorIndex = ordered(0)
                    If anchorIndex >= 0 AndAlso anchorIndex < ListView1.Items.Count Then
                        gCurrentlySelectedListViewItemIndex = anchorIndex
                    Else
                        gCurrentlySelectedListViewItemIndex = ordered(0)
                    End If
                End If
            End If

            If ListView1.SelectedItems.Count = 0 AndAlso fallbackToFirst AndAlso ListView1.Items.Count > 0 Then
                ListView1.SelectedItems.Add(ListView1.Items(0))
                gCurrentlySelectedListViewItemIndex = 0
            End If

            If ListView1.SelectedItems.Count > 0 Then
                gPendingSelectionApplied = True
            End If

            'Ensure selection is scrolled into view and focus is placed on the selected container.
            If ListView1.SelectedItems.Count > 0 Then
                ListView1.UpdateLayout()
                ListView1.ScrollIntoView(ListView1.SelectedItems(0))
            End If

            'Refresh detail view to reflect the restored selection.
            UpdateDetails()

        Finally

            'Reattach selection handler now that restore is complete.
            AddHandler ListView1.SelectionChanged, AddressOf ListView1_SelectionChanged

        End Try


    End Sub

    Private Sub SelectAllMembersOfAnEmailChain()

        Try
            If ListView1.SelectedItems.Count = 0 Then
                Exit Sub
            End If

            'work with most recent selected item
            Dim SelectedListViewItem = TryCast(
            ListView1.SelectedItems.Item(ListView1.SelectedItems.Count - 1),
            ListViewRowClass)

            If SelectedListViewItem Is Nothing Then
                Exit Sub
            End If

            UpdateDetails()

            If SelectedListViewItem.ChainIndicator = ListViewRowClass.ChainIndicatorValues.NotPartOfAChain Then
                ' nothing to do
            Else
                Dim PrevSubjectTrailer As String = ""
                Dim CurrentSubjectPlusTrailer As String = SelectedListViewItem.Subject & SelectedListViewItem.Trailer

                'find top of chain and select it
                Dim i As Integer
                For i = SelectedListViewItem.Index To 0 Step -1
                    Dim item = TryCast(ListView1.Items(i), ListViewRowClass)
                    If item Is Nothing Then Exit For

                    PrevSubjectTrailer = item.Subject & item.Trailer
                    If CurrentSubjectPlusTrailer <> PrevSubjectTrailer Then
                        i += 1
                        SelectListViewItem(i)
                        Exit For
                    End If
                    If i = 0 Then
                        SelectListViewItem(i)
                        Exit For
                    End If
                Next

                'Select remaining emails in the same chain
                For i = i To ListView1.Items.Count - 2
                    Dim item = TryCast(ListView1.Items(i), ListViewRowClass)
                    Dim nextItem = TryCast(ListView1.Items(i + 1), ListViewRowClass)
                    If item Is Nothing OrElse nextItem Is Nothing Then Exit For

                    If (item.Subject = nextItem.Subject) AndAlso (item.Trailer = nextItem.Trailer) Then
                        SelectListViewItem(i)
                    Else
                        Exit For
                    End If

                    If i < ListView1.Items.Count - 1 Then
                        SelectListViewItem(i + 1)
                    End If
                Next
            End If

        Catch ex As Exception
            ' optionally log ex
        End Try

        UpdateMainMessageLine()

    End Sub

    Private Sub SelectListViewItem(ByVal index As Integer)

        Dim EntryHasAlreadyBeenSelected As Boolean = False

        For i = 0 To ListView1.SelectedItems.Count - 1
            If ListView1.SelectedItems(i).index = index Then
                EntryHasAlreadyBeenSelected = True
                Exit For
            End If
        Next

        If EntryHasAlreadyBeenSelected Then
        Else
            ListView1.SelectedItems.Add(ListView1.Items(index))
        End If

    End Sub

    Private Sub UpdateMainMessageLine()

        Try

            Select Case ListView1.Items.Count
                Case Is = 0
                    Me.lblMainMessageLine.Content = "0 e-mails"
                Case Is = 1
                    Me.lblMainMessageLine.Content = "1 e-mail"
                Case Else
                    Me.lblMainMessageLine.Content = ListView1.Items.Count.ToString("#,#", System.Globalization.CultureInfo.InvariantCulture) & " e-mails"
            End Select

            If ListView1.SelectedItems.Count = 0 Then
                Me.lblMainMessageLine.Content &= " (0 selected)"
            Else
                Me.lblMainMessageLine.Content &= " (" & ListView1.SelectedItems.Count.ToString("#,#", System.Globalization.CultureInfo.InvariantCulture) & " selected)"
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub UpdateDetails()

        Try

            BlankOutDetails()

            Dim selected = TryCast(ListView1.SelectedItem, ListViewRowClass)
            If selected Is Nothing Then
                Exit Try
            End If

            With selected

                Me.tbDetailSubject.Text = .Subject
                Me.tbDetailFrom.Text = .From
                Me.tbDetailTo.Text = .xTo
                Me.tbDetailDateTime.Text = Format(.DateTime, gPreferredDateFormat & " " & gPreferredTimeFormat)
                Me.tbDetailOrginal.Text = gFolderNamesTable(.OriginalFolder).TrimStart("\"c)

                ' Only offer pick folders when there is a valid recommended folder (RecommendedFolderFinal >= 0)
                Dim hasValidRecommendation As Boolean = (.RecommendedFolderFinal >= 0)

                If hasValidRecommendation Then
                    Me.tbDetailTarget1.Text = gFolderNamesTable(.RecommendedFolderFinal)
                Else
                    Me.tbDetailTarget1.Text = ""
                End If

                If gPickAFolderWindow IsNot Nothing Then
                    If hasValidRecommendation Then
                        gPickAFolderWindow.intRecommendation1 = .RecommendedFolder1
                        gPickAFolderWindow.intRecommendation2 = .RecommendedFolder2
                        gPickAFolderWindow.intRecommendation3 = .RecommendedFolder3
                        gPickAFolderWindow.intRecommendation4 = .RecommendedFolderFinal
                    Else
                        ' No valid picks – ensure Pick A Folder window offers none
                        gPickAFolderWindow.intRecommendation1 = -1
                        gPickAFolderWindow.intRecommendation2 = -1
                        gPickAFolderWindow.intRecommendation3 = -1
                        gPickAFolderWindow.intRecommendation4 = -1
                    End If

                    gPickAFolderWindow.UpdateRecommendationsOnPickAFolderWindow()
                End If
            End With

        Catch ex As Exception

        Finally
            UpdateMainMessageLine()
        End Try

    End Sub

    Private Sub BlankOutDetails()

        Me.tbDetailSubject.Text = ""
        Me.tbDetailFrom.Text = ""
        Me.tbDetailTo.Text = ""
        Me.tbDetailDateTime.Text = ""
        Me.tbDetailOrginal.Text = ""
        Me.tbDetailTarget1.Text = ""

    End Sub

    Private Sub MenuActions_Click(ByVal sender As System.Object,
                              ByVal e As System.Windows.RoutedEventArgs) Handles _
    MenuOpen.Click, MenuHide.Click, MenuDelete.Click, MenuExit.Click,
    MenuViewRead.Click, MenuViewUnRead.Click,
    MenuViewAll.Click, MenuViewInbox.Click, MenuViewSent.Click,
    MenuUndo.Click, MenuHelpSub.Click, MenuAbout.Click, MenuOptions.Click, MenuRefresh.Click, MenuQuickRefresh.Click,
    MenuContextDelete.Click, MenuContextHide.Click, MenuContextOpen.Click,
    MenuContextToggleRead.Click

        If sender.GetType.ToString = "System.Windows.Controls.Button" Then
            PerformAction(sender.Tag) ' for buttons
        Else
            PerformAction(sender.Tag, sender.IsChecked) ' for menu items
        End If

    End Sub


    Private Sub FileMenuActions_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles _
               MenuContextFile1.Click, MenuContextFile2.Click, MenuContextFile3.Click, MenuContextFile4.Click


        Dim SelectedFolder As String = ""

        Select Case sender.name
            Case Is = "MenuContextFile1"
                SelectedFolder = Me.MenuContextFile1.Header
            Case Is = "MenuContextFile2"
                SelectedFolder = Me.MenuContextFile2.Header
            Case Is = "MenuContextFile3"
                SelectedFolder = Me.MenuContextFile3.Header
            Case Is = "MenuContextFile4"
                SelectedFolder = Me.MenuContextFile4.Header
        End Select

        SelectedFolder = SelectedFolder.Remove(0, 9)

        gPickFromContextMenuOverride = LookupFolderNamesTableIndex(SelectedFolder)

        PerformAction(sender.tag, sender.ischecked)

        gPickFromContextMenuOverride = -1

    End Sub

    ' Helper: extract mailbox/postbox name based on Outlook store display name
    Private Function GetMailboxNameFromFolderPath(ByVal folderPath As String, ByVal storeId As String) As String
        Try
            If oNS IsNot Nothing AndAlso Not String.IsNullOrEmpty(storeId) Then
                Dim store As Microsoft.Office.Interop.Outlook.Store = oNS.GetStoreFromID(storeId)
                If store IsNot Nothing AndAlso store.DisplayName IsNot Nothing Then
                    Return store.DisplayName
                End If
            End If
        Catch
            ' fall back to folderPath parsing below
        End Try

        If String.IsNullOrEmpty(folderPath) Then Return ""
        Dim parts As String() = folderPath.Split("\"c)
        If parts.Length > 1 Then
            Return parts(1)
        End If
        Return folderPath.Trim("\"c)
    End Function

    Public Sub SafelyActivateMenu()
        Call Dispatcher.BeginInvoke(ActivateMenu)
    End Sub
    Private ActivateMenu As New System.Windows.Forms.MethodInvoker(AddressOf ActivateMenuNow)
    Private Sub ActivateMenuNow()

        Me.Menu1.Focus()

    End Sub

    Public Sub SafelyPerformActionByProxy()
        Call Dispatcher.BeginInvoke(PerformActionByProxy)
    End Sub
    Private PerformActionByProxy As New System.Windows.Forms.MethodInvoker(AddressOf PerformActionByProxyNow)
    Private Sub PerformActionByProxyNow()

        PerformAction(gProxyAction)

    End Sub

    Enum FirstSelectionState
        Unknown
        ReadToUnread
        UnreadToRead
    End Enum

    Private Sub ToggleReadStateForSelectedItem()

        Try

            ' Show wait cursor while we do COM work
            SetMousePointer(Cursors.Wait)

            If ListView1.SelectedItems.Count = 0 Then
                Exit Try
            End If

            ' Ensure Outlook is running and session is usable
            If Not EnsureOutlookIsRunning() Then
                Exit Try
            End If

            RemoveHandler ListView1.SelectionChanged, AddressOf ListView1_SelectionChanged

            Dim selectionSnapshot As SelectionSnapshot = CaptureSelectionSnapshot()
            Dim indicesToUpdate As New List(Of Integer)()
            Dim errorOccurred As Boolean = False

            gSuppressUpdatesToDetailBox = True

            ' if multiple items are selected, we can't just toggle each one
            ' the user will be prompted to "Mark as Read" or "Mark as Unread" based on the first selected item
            ' so rather than simply toggling the read/unread state in the UI
            ' we need to figure out what the first entry will toggle to and then toggle all selected items to that same state
            ' (all in the name of usability and consistency)

            Dim firstSelection As FirstSelectionState = FirstSelectionState.Unknown

            Try
                ' Process all selected items
                For Each selectedItem In ListView1.SelectedItems
                    Dim index As Integer = selectedItem.Index
                    Dim selectedRow As ListViewRowClass = TryCast(ListView1.Items(index), ListViewRowClass)

                    If selectedRow Is Nothing Then
                        Continue For
                    End If

                    Dim entryId As String = selectedRow.OutlookEntryID
                    If String.IsNullOrEmpty(entryId) Then
                        Continue For
                    End If

                    Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = Nothing

                    ' We will try GetItemFromID up to 2 times:
                    '  - first with the current session
                    '  - on RPC disconnect errors, rebuild session and retry once
                    Dim attempt As Integer = 0
                    While attempt < 2 AndAlso mailItem Is Nothing

                        Try
                            mailItem = TryCast(oNS.GetItemFromID(entryId), Microsoft.Office.Interop.Outlook.MailItem)

                        Catch comEx As System.Runtime.InteropServices.COMException
                            Const RPC_E_SERVER_UNAVAILABLE As Integer = &H800706BA
                            Const RPC_E_DISCONNECTED As Integer = &H800706BE

                            If (comEx.HResult = RPC_E_SERVER_UNAVAILABLE OrElse comEx.HResult = RPC_E_DISCONNECTED) AndAlso attempt = 0 Then
                                ' Drop and rebuild the Outlook session once
                                oNS = Nothing
                                oApp = Nothing

                                If Not EnsureOutlookIsRunning() Then
                                    errorOccurred = True
                                    Exit While
                                End If

                                ' Let the loop retry GetItemFromID with fresh session
                                mailItem = Nothing
                            Else
                                ' Any other COM error, or second failure: stop processing
                                errorOccurred = True
                                Exit While
                            End If

                        Catch ex As Exception
                            errorOccurred = True
                            Exit While
                        End Try

                        attempt += 1
                    End While

                    If mailItem Is Nothing Then
                        ' Could not get mail item, skip this one
                        Continue For
                    End If

                    If firstSelection = FirstSelectionState.Unknown Then
                        If mailItem.UnRead Then
                            firstSelection = FirstSelectionState.UnreadToRead
                        Else
                            firstSelection = FirstSelectionState.ReadToUnread
                        End If
                    End If

                    ' set the read/unread flag in Outlook
                    Try

                        If mailItem.UnRead Then

                            If firstSelection = FirstSelectionState.UnreadToRead Then
                                ' Mark as read
                                mailItem.UnRead = False
                            End If
                        Else

                            If firstSelection = FirstSelectionState.ReadToUnread Then
                                ' Mark as unread
                                mailItem.UnRead = True
                            End If

                        End If

                        Dim action As String = If(mailItem.UnRead, "ReadGoingToUnread", "UnreadGoingToRead")
                        Call _MainWindow.BlockDuplicateEventProcessing(action, mailItem.EntryID)


                        mailItem.Save()
                        indicesToUpdate.Add(index)
                    Catch comEx As System.Runtime.InteropServices.COMException
                        Const RPC_E_SERVER_UNAVAILABLE As Integer = &H800706BA
                        Const RPC_E_DISCONNECTED As Integer = &H800706BE

                        If comEx.HResult = RPC_E_SERVER_UNAVAILABLE OrElse comEx.HResult = RPC_E_DISCONNECTED Then
                            errorOccurred = True
                            Exit For
                        Else
                            ' Skip this item and continue with others
                            Continue For
                        End If
                    Catch ex As Exception
                        ' Skip this item and continue with others
                        Continue For
                    End Try
                Next


                ' Update UI for all successfully toggled items
                For Each index In indicesToUpdate

                    Dim updatedRow As ListViewRowClass = CType(ListView1.Items(index), ListViewRowClass)

                    ' Get the mail item again to get the updated state
                    Dim entryId As String = updatedRow.OutlookEntryID

                    If Not String.IsNullOrEmpty(entryId) Then

                        Try
                            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(oNS.GetItemFromID(entryId), Microsoft.Office.Interop.Outlook.MailItem)
                            If mailItem IsNot Nothing Then
                                If mailItem.UnRead Then
                                    updatedRow.UnRead = System.Windows.FontWeights.Bold
                                Else
                                    updatedRow.UnRead = System.Windows.FontWeights.Normal
                                End If
                            End If
                        Catch
                            ' Skip if we can't get the item
                        End Try
                    End If

                    ' Replace item in the ListView so the binding updates
                    ListView1.Items.RemoveAt(index)
                    ListView1.Items.Insert(index, updatedRow)
                Next

                ' Restore selections and focus

                ListView1.UpdateLayout()
                If indicesToUpdate.Count > 0 Then
                    ListView1.SelectedIndex = indicesToUpdate(0)
                    ListView1.ScrollIntoView(ListView1.Items(indicesToUpdate(0)))
                End If

                ' Also update details pane and counters
                UpdateDetails()

                If errorOccurred Then

                    Call ShowMessageBox("Toggle Read/Unread Failed",
                         CustomDialog.CustomDialogIcons.Information,
                         "FileFriendly could not update the read/unread state for all selected items.",
                         "Some items may have been successfully toggled, but not all.",
                         "ToggleReadStateForSelectedItem()",
                         "",
                         CustomDialog.CustomDialogIcons.None,
                         CustomDialog.CustomDialogButtons.OK,
                         CustomDialog.CustomDialogResults.OK)

                End If

            Finally
                gSuppressUpdatesToDetailBox = False
            End Try

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Toggle Read/Unread Failed",
                 CustomDialog.CustomDialogIcons.Stop,
                 "FileFriendly could not update the read/unread state for all selected items.",
                 "Some items may have been successfully toggled, but not all.",
                 ex.TargetSite.Name & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try

        ' Always restore cursor
        AddHandler ListView1.SelectionChanged, AddressOf ListView1_SelectionChanged
        SetMousePointer(Cursors.Arrow)

    End Sub

    Public Sub SafelyUpdateReadToggleContextMenu()
        Call Dispatcher.BeginInvoke(UpdateReadToggleContextMenu)
    End Sub

    Private UpdateReadToggleContextMenu As New System.Windows.Forms.MethodInvoker(AddressOf UpdateReadToggleContextMenuNow)

    Private Sub UpdateReadToggleContextMenuNow()

        Try

            Dim row As ListViewRowClass = TryCast(ListView1.SelectedItem, ListViewRowClass)
            If row Is Nothing Then
                Me.MenuContextToggleRead.Visibility = Windows.Visibility.Collapsed
                Exit Sub
            End If

            ' Bold = unread, Normal = read
            Dim isUnread As Boolean = (row.UnRead = System.Windows.FontWeights.Bold)

            If isUnread Then
                Me.MenuContextToggleRead.Header = "Mark as Read"
            Else
                Me.MenuContextToggleRead.Header = "Mark as Unread"
            End If

            Me.MenuContextToggleRead.Visibility = Windows.Visibility.Visible

        Catch ex As Exception
            ' Keep context‑menu failures silent, consistent with rest of file
        End Try

    End Sub

    Private Sub StartRefresh(ByVal QuickRefresh As Boolean)

        gRefreshConfirmed = False

        Try
            If gIsRefreshing Then
                gCancelRefresh = True
                Exit Sub
            End If

            If ActionLogIndex > 0 Then

                If ShowMessageBox("FileFriendly",
                       CustomDialog.CustomDialogIcons.Question,
                       "Please note:",
                       "If you refresh you will no longer be able to undo the changes you have made up until now." & vbCrLf & vbCrLf &
                       "Would you still like to refresh?",
                       "You will however be able to undo future changes.",
                       "",
                       CustomDialog.CustomDialogIcons.None,
                       CustomDialog.CustomDialogButtons.YesNo,
                       CustomDialog.CustomDialogResults.Yes) = CustomDialog.CustomDialogResults.No Then
                    Exit Sub
                End If

            End If

            MenuOptionEnabled("Undo", False)
            ActionLogIndex = 0

            If QuickRefresh Then
                gRefreshConfirmed = True
            Else
                gPickARefreshModeWindow = New PickARefreshMode
                gPickARefreshModeWindow.ShowDialog()
                gPickARefreshModeWindow = Nothing
            End If

            If gRefreshConfirmed Then

                If gRefreshInbox Or gRefreshSent Or gRefreshOtherFolders Then
                    MenuRefresh.Foreground = gForegroundColourEnabled
                    MenuActions.Foreground = gForegroundColourEnabled
                    RefreshGrid(False, False, QuickRefresh)
                Else

                    Call ShowMessageBox("FileFriendly",
                         CustomDialog.CustomDialogIcons.Warning,
                         "Note!",
                         "Inbox, sent items and other folders shouldn`t all be unchecked at the same time.",
                         "If you unchecked all three then there will be nothing to review!",
                         "",
                         CustomDialog.CustomDialogIcons.None,
                         CustomDialog.CustomDialogButtons.OK,
                         CustomDialog.CustomDialogResults.OK)

                    ClearGrid()
                End If

            End If

        Catch
        End Try
    End Sub

    Private Sub PerformAction(ByVal Action As String, Optional ByRef flag As Boolean = True)

        MenuKeyStrokeOverRide = False

        Me.Cursor = Cursors.Wait

        Try

            Select Case Action

                Case Is = "Open"

                    If My.Settings.ConfirmOpen Then
                        If ConfirmActionMessage(Action) Then OpenAnEmail()
                    Else
                        OpenAnEmail()
                    End If

                Case Is = "File", "Delete", "Hide"

                    If ConfirmActionMessage(Action) Then
                        ActionRequestAgainstAllSelectedItems(Action, Me.ListView1)
                    End If

                Case Is = "ToggleRead"

                    ToggleReadStateForSelectedItem()

                Case Is = "Options"

                    gARefreshIsRequired = False

                    gOptionsWindow = New OptionsWindow
                    gOptionsWindow.ShowDialog()
                    gOptionsWindow = Nothing

                    gRefreshInbox = My.Settings.ScanInbox
                    gRefreshSent = My.Settings.ScanSent
                    gRefreshOtherFolders = My.Settings.ScanAll
                    gAutoChainSelect = My.Settings.AutoChainSelect

                    'If gARefreshIsRequired Then
                    '    RefreshGrid(False, False, True)
                    'End If

                Case Is = "Undo"

                    If My.Settings.ConfirmUndo Then
                        If ConfirmActionMessage(Action) Then UndoLastAction()
                    Else
                        UndoLastAction()
                    End If

                Case Is = "Quick Refresh [F5]"

                    StartRefresh(True)

                Case Is = "Refresh"

                    StartRefresh(False)

                Case Is = "Exit"

                    If My.Settings.ConfirmExit Then
                        If ConfirmActionMessage(Action) Then ShutDown()
                    Else
                        ShutDown()
                    End If

                Case Is = "ViewInbox"
                    gViewInbox = flag
                    ValidateInboxSentFoldersCombinatation()
                    ApplyFilter()

                Case Is = "ViewSent"
                    gViewSent = flag
                    ValidateInboxSentFoldersCombinatation()
                    ApplyFilter()

                Case Is = "ViewAll"
                    gViewAll = flag
                    ValidateInboxSentFoldersCombinatation()
                    ApplyFilter()

                Case Is = "ViewRead"
                    gViewRead = flag
                    ValidateReadUnReadCombinatation()
                    ApplyFilter()

                Case Is = "ViewUnRead"
                    gViewUnRead = flag
                    ValidateReadUnReadCombinatation()
                    ApplyFilter()

                Case Is = "Help"
                    System.Diagnostics.Process.Start(gHelpWebPage)
                    System.Threading.Thread.Sleep(3000)

                Case Is = "About"

                    gAboutWindow = New LicenseWindow
                    gAboutWindow.ShowDialog()
                    gAboutWindow = Nothing

            End Select

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Perform Action Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 ex.TargetSite.Name & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try

        Me.Cursor = Cursors.Arrow

    End Sub

    Private Sub ValidateInboxSentFoldersCombinatation()

        If gViewInbox Or gViewSent Or gViewAll Or (Me.MenuRefresh.Foreground Is gForegroundColourAlert) Then

            If Me.MenuViewInbox.IsEnabled Then
                Me.MenuViewInbox.Foreground = gForegroundColourEnabled
            Else
                Me.MenuViewInbox.Foreground = gForegroundColourDisabled
            End If

            If Me.MenuViewSent.IsEnabled Then
                Me.MenuViewSent.Foreground = gForegroundColourEnabled
            Else
                Me.MenuViewSent.Foreground = gForegroundColourDisabled
            End If

            If Me.MenuViewAll.IsEnabled Then
                Me.MenuViewAll.Foreground = gForegroundColourEnabled
            Else
                Me.MenuViewAll.Foreground = gForegroundColourDisabled
            End If

        Else

            Call ShowMessageBox("FileFriendly",
                     CustomDialog.CustomDialogIcons.Warning,
                     "Note!",
                     "Inbox, sent items and other folders shouldn`t all be unchecked at the same time.",
                     "If you unchecked all three then there will be nothing to review!",
                     "",
                     CustomDialog.CustomDialogIcons.None,
                     CustomDialog.CustomDialogButtons.OK,
                     CustomDialog.CustomDialogResults.OK)

            If Me.MenuViewInbox.IsEnabled Then
                Me.MenuViewInbox.Foreground = gForegroundColourAlert
            Else
                Me.MenuViewInbox.Foreground = gForegroundColourDisabled
            End If

            If Me.MenuViewSent.IsEnabled Then
                Me.MenuViewSent.Foreground = gForegroundColourAlert
            Else
                Me.MenuViewSent.Foreground = gForegroundColourDisabled
            End If

            If Me.MenuViewAll.IsEnabled Then
                Me.MenuViewAll.Foreground = gForegroundColourAlert
            Else
                Me.MenuViewAll.Foreground = gForegroundColourDisabled
            End If

        End If

        If (Me.MenuViewRead.Foreground Is gForegroundColourAlert) Or (Me.MenuViewInbox.Foreground Is gForegroundColourAlert) Then
            Me.MenuView.Foreground = gForegroundColourAlert
        Else
            Me.MenuView.Foreground = gForegroundColourEnabled
        End If

    End Sub

    Private Sub ValidateReadUnReadCombinatation()

        If gViewRead Or gViewUnRead Or (Me.MenuRefresh.Foreground Is gForegroundColourAlert) Then

            Me.MenuViewRead.Foreground = gForegroundColourEnabled
            Me.MenuViewUnRead.Foreground = gForegroundColourEnabled

        Else

            Call ShowMessageBox("FileFriendly",
                     CustomDialog.CustomDialogIcons.Warning,
                     "Note!",
                     "Read and Unread shouldn`t both be unchecked at the same time.",
                     "If you unchecked them both then there will be nothing to review!",
                     "",
                     CustomDialog.CustomDialogIcons.None,
                     CustomDialog.CustomDialogButtons.OK,
                     CustomDialog.CustomDialogResults.OK)

            Me.MenuViewRead.Foreground = gForegroundColourAlert
            Me.MenuViewUnRead.Foreground = gForegroundColourAlert

        End If

        If (Me.MenuViewRead.Foreground Is gForegroundColourAlert) Or (Me.MenuViewInbox.Foreground Is gForegroundColourAlert) Then
            Me.MenuView.Foreground = gForegroundColourAlert
        Else
            Me.MenuView.Foreground = gForegroundColourEnabled
        End If

    End Sub

    Private Sub ShutDown()

        gClosingNow = True
        Me.Visibility = Windows.Visibility.Hidden

        If gPickAFolderWindow IsNot Nothing Then
            gPickAFolderWindow.Visibility = Windows.Visibility.Hidden
        End If

        Me.Close()

    End Sub
    Private Sub OpenAnEmail()

        ' Show wait cursor while we do COM work
        SetMousePointer(Cursors.Wait)

        Try

            Dim selectedRow As ListViewRowClass = TryCast(ListView1.SelectedItem, ListViewRowClass)
            If selectedRow Is Nothing OrElse
               String.IsNullOrEmpty(selectedRow.OutlookEntryID) Then
                Exit Try
            End If

            ' Ensure Outlook is running and session is usable
            If Not EnsureOutlookIsRunning() Then
                Exit Try
            End If

            Dim entryId As String = selectedRow.OutlookEntryID
            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = Nothing

            ' Retry GetItemFromID once if we see an RPC disconnect
            Dim attempt As Integer = 0
            While attempt < 2 AndAlso mailItem Is Nothing

                Try
                    mailItem = TryCast(oNS.GetItemFromID(entryId), Microsoft.Office.Interop.Outlook.MailItem)

                Catch comEx As System.Runtime.InteropServices.COMException
                    Const RPC_E_SERVER_UNAVAILABLE As Integer = &H800706BA
                    Const RPC_E_DISCONNECTED As Integer = &H800706BE

                    If (comEx.HResult = RPC_E_SERVER_UNAVAILABLE OrElse comEx.HResult = RPC_E_DISCONNECTED) AndAlso attempt = 0 Then
                        ' Drop and rebuild the Outlook session once
                        oNS = Nothing
                        oApp = Nothing

                        If Not EnsureOutlookIsRunning() Then
                            Exit Try
                        End If

                        ' Loop will retry with fresh session
                        mailItem = Nothing
                    Else

                        Call ShowMessageBox("FileFriendly - Open An E-mail Fail",
                             CustomDialog.CustomDialogIcons.Stop,
                             "FileFriendly could not open the selected e-mail",
                             "FileFriendly could not open the selected e-mail in Outlook. (1)",
                             comEx.TargetSite.Name & " - " & comEx.ToString,
                             "",
                             CustomDialog.CustomDialogIcons.None,
                             CustomDialog.CustomDialogButtons.OK,
                             CustomDialog.CustomDialogResults.OK)

                        Exit Try
                    End If

                Catch ex As Exception

                    Call ShowMessageBox("FileFriendly - Open An E-mail Fail",
                         CustomDialog.CustomDialogIcons.Stop,
                         "FileFriendly could not open the selected e-mail",
                         "FileFriendly could not open the selected e-mail in Outlook. (2)",
                         ex.TargetSite.Name & " - " & ex.ToString,
                         "",
                         CustomDialog.CustomDialogIcons.None,
                         CustomDialog.CustomDialogButtons.OK,
                         CustomDialog.CustomDialogResults.OK)

                    Exit Try

                End Try

                attempt += 1
            End While

            If mailItem Is Nothing Then
                Exit Try
            End If

            mailItem.Display()

            ' if the email isn't already marked then mark it read in the grid
            Dim index As Integer = ListView1.SelectedIndex
            If index >= 0 AndAlso ListView1.Items(index).UnRead = System.Windows.FontWeights.Bold Then
                Dim hold As ListViewRowClass = CType(ListView1.Items(index), ListViewRowClass)
                hold.UnRead = System.Windows.FontWeights.Normal
                ListView1.Items.RemoveAt(index)
                ListView1.Items.Insert(index, hold)
                ListView1.SelectedIndex = index
                UpdateDetails()
            End If

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Open An E-mail Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error." & vbCrLf & "If Outlook is not running please start it and try again",
                 ex.TargetSite.Name & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        Finally
            SetMousePointer(Cursors.Arrow)
        End Try

    End Sub

    Private Sub UndoLastAction()

        Try

            If ActionLogIndex < 1 Then
                MenuOptionEnabled("Undo", False)
                ActionLogIndex = 0
                Exit Try
            End If

            Dim NewOutlookEntryID As String = ""

            'point back to the last populated log entryId
            ActionLogIndex -= 1

            Dim i As Integer = 0
            While ActionLog(ActionLogIndex, i).ActionApplied > 0

                If (ActionLog(ActionLogIndex, i).ActionApplied = ActionType.File) Or
                   (ActionLog(ActionLogIndex, i).ActionApplied = ActionType.Delete) Then

                    'reverse file all sub item actions

                    NewOutlookEntryID =
                               FileMessage("Undo",
                               ActionLog(ActionLogIndex, i).FixedIndex,
                               ActionLog(ActionLogIndex, i).EmailID,
                               Nothing,
                               ActionLog(ActionLogIndex, i).SourceStoreID,
                               ActionLog(ActionLogIndex, i).TargetEntryID,
                               ActionLog(ActionLogIndex, i).TargetStoreID)

                    gFinalRecommendationTable(ActionLog(ActionLogIndex, i).FixedIndex).OutlookEntryID = NewOutlookEntryID

                Else

                    NewOutlookEntryID = gFinalRecommendationTable(ActionLog(ActionLogIndex, i).FixedIndex).OutlookEntryID

                End If

                ' re-establish the list view item:
                ' changing to gFinalRecommendationTable(-).index to a value > -1 
                ' effectively un-deletes the list view item once the list view is refreshed
                gFinalRecommendationTable(ActionLog(ActionLogIndex, i).FixedIndex).Index = 1

                ActionLog(ActionLogIndex, i).ActionApplied = Nothing
                ActionLog(ActionLogIndex, i).FixedIndex = Nothing
                ActionLog(ActionLogIndex, i).EmailID = Nothing
                ActionLog(ActionLogIndex, i).SourceStoreID = Nothing
                ActionLog(ActionLogIndex, i).TargetEntryID = Nothing
                ActionLog(ActionLogIndex, i).TargetStoreID = Nothing

                i += 1

            End While

            ApplyFilter() 'force the list view to be rebuilt, adding back in any undone items

            'determine which entryId should be the new selected entryId 
            Dim IndexToBePositionedAt As Integer = 0
            If ListView1.Items.Count > 0 Then
                For ii As Integer = 0 To ListView1.Items.Count - 1
                    If ListView1.Items(ii).OutlookEntryID = NewOutlookEntryID Then
                        IndexToBePositionedAt = ii
                        Exit For
                    End If
                Next
            End If

            ListView1.SelectedItem = ListView1.Items(IndexToBePositionedAt)
            If gAutoChainSelect Then SelectAllMembersOfAnEmailChain()
            ListView1.Focus()

            UpdateDetails()

            If ActionLogIndex = 0 Then
                MenuOptionEnabled("Undo", False)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub MenuOptionEnabled(ByVal strOption As String, ByVal flag As Boolean)

        Select Case strOption

            Case Is = "Undo"
                If flag Then
                    Me.MenuUndo.Foreground = gForegroundColourEnabled
                Else
                    Me.MenuUndo.Foreground = gForegroundColourDisabled
                End If
                Me.MenuUndo.IsEnabled = flag

            Case Is = "Refresh"
                If flag Then
                    Me.MenuRefresh.Foreground = gForegroundColourEnabled
                    Me.MenuQuickRefresh.Foreground = gForegroundColourEnabled
                Else
                    Me.MenuRefresh.Foreground = gForegroundColourDisabled
                    Me.MenuQuickRefresh.Foreground = gForegroundColourDisabled
                End If
                Me.MenuRefresh.IsEnabled = flag

            Case Is = "Options"
                If flag Then
                    Me.MenuOptions.Foreground = gForegroundColourEnabled
                Else
                    Me.MenuOptions.Foreground = gForegroundColourDisabled
                End If
                Me.MenuOptions.IsEnabled = flag

            Case Is = "Hide"
                If flag Then
                    Me.MenuHide.Foreground = gForegroundColourEnabled
                Else
                    Me.MenuHide.Foreground = gForegroundColourDisabled
                End If
                Me.MenuHide.IsEnabled = flag

            Case Is = "Open"
                If flag Then
                    Me.MenuOpen.Foreground = gForegroundColourEnabled
                Else
                    Me.MenuOpen.Foreground = gForegroundColourDisabled
                End If
                Me.MenuOpen.IsEnabled = flag

            Case Is = "Delete"
                If flag Then
                    Me.MenuDelete.Foreground = gForegroundColourEnabled
                Else
                    Me.MenuDelete.Foreground = gForegroundColourDisabled
                End If
                Me.MenuDelete.IsEnabled = flag

            Case Is = "View"

                If flag Then
                    Me.MenuViewInbox.Foreground = gForegroundColourEnabled
                    Me.MenuViewSent.Foreground = gForegroundColourEnabled
                    Me.MenuViewAll.Foreground = gForegroundColourEnabled
                    Me.MenuViewRead.Foreground = gForegroundColourEnabled
                    Me.MenuViewUnRead.Foreground = gForegroundColourEnabled
                Else
                    Me.MenuViewInbox.Foreground = gForegroundColourDisabled
                    Me.MenuViewSent.Foreground = gForegroundColourDisabled
                    Me.MenuViewAll.Foreground = gForegroundColourDisabled
                    Me.MenuViewRead.Foreground = gForegroundColourDisabled
                    Me.MenuViewUnRead.Foreground = gForegroundColourDisabled
                End If
                Me.MenuViewInbox.IsEnabled = flag
                Me.MenuViewSent.IsEnabled = flag
                Me.MenuViewAll.IsEnabled = flag
                Me.MenuViewRead.IsEnabled = flag
                Me.MenuViewUnRead.IsEnabled = flag

                If gRefreshInbox Then
                Else
                    Me.MenuViewInbox.IsEnabled = False
                    Me.MenuViewInbox.Foreground = gForegroundColourDisabled
                End If

                If gRefreshSent Then
                Else
                    Me.MenuViewSent.IsEnabled = False
                    Me.MenuViewSent.Foreground = gForegroundColourDisabled
                End If

                If gRefreshOtherFolders Then
                Else
                    Me.MenuViewAll.IsEnabled = False
                    Me.MenuViewAll.Foreground = gForegroundColourDisabled
                End If

        End Select

    End Sub

    Public Function FileMessage(ByVal Action As String,
                            ByVal FixedIndex As Integer,
                            ByVal EmailID As String,
                            ByVal SourceEntryID As String,
                            ByVal SourceStoreID As String,
                            ByVal TargetEntryID As String,
                            ByVal TargetStoreID As String) As String

        'Returns new email id of filed message (empty string on failure)

        Dim ReturnCode As String = ""

        ' Show wait cursor during Outlook COM work
        SetMousePointer(Cursors.Wait)

        Try
            If Not EnsureOutlookIsRunning() Then
                Return ""
            End If

            Dim mail As Microsoft.Office.Interop.Outlook.MailItem = Nothing
            Dim targetFolder As Microsoft.Office.Interop.Outlook.MAPIFolder = Nothing

            ' Retry GetItemFromID / GetFolderFromID once on RPC disconnect
            Dim attempt As Integer = 0
            While attempt < 2 AndAlso (mail Is Nothing OrElse targetFolder Is Nothing)

                Try
                    mail = TryCast(oNS.GetItemFromID(EmailID, SourceStoreID), Microsoft.Office.Interop.Outlook.MailItem)
                    targetFolder = oNS.GetFolderFromID(TargetEntryID, TargetStoreID)

                Catch comEx As System.Runtime.InteropServices.COMException
                    Const RPC_E_SERVER_UNAVAILABLE As Integer = &H800706BA
                    Const RPC_E_DISCONNECTED As Integer = &H800706BE

                    If (comEx.HResult = RPC_E_SERVER_UNAVAILABLE OrElse comEx.HResult = RPC_E_DISCONNECTED) AndAlso attempt = 0 Then
                        oNS = Nothing
                        oApp = Nothing

                        If Not EnsureOutlookIsRunning() Then
                            Return ""
                        End If

                        mail = Nothing
                        targetFolder = Nothing
                    Else

                        Call ShowMessageBox("FileFriendly - Outlook Error",
                             CustomDialog.CustomDialogIcons.Information,
                             "Unexpected Error!",
                             "FileFriendly has encountered an unexpected error." & vbCrLf & "FileFriendly could not complete the requested action in Outlook. (1)",
                             comEx.TargetSite.Name & " - " & comEx.ToString,
                             "",
                             CustomDialog.CustomDialogIcons.None,
                             CustomDialog.CustomDialogButtons.OK,
                             CustomDialog.CustomDialogResults.OK)

                        Return ""
                    End If

                Catch ex As Exception

                    Call ShowMessageBox("FileFriendly - Outlook Error",
                         CustomDialog.CustomDialogIcons.Information,
                         "Unexpected Error!",
                         "FileFriendly has encountered an unexpected error." & vbCrLf & "FileFriendly could not complete the requested action in Outlook. (2)",
                         ex.TargetSite.Name & " - " & ex.ToString,
                         "",
                         CustomDialog.CustomDialogIcons.None,
                         CustomDialog.CustomDialogButtons.OK,
                         CustomDialog.CustomDialogResults.OK)

                    Return ""
                End Try

                attempt += 1
            End While

            If mail Is Nothing OrElse targetFolder Is Nothing Then
                Return ""
            End If

            ' the move being done below will itself raise a 'Remove' event that we need to ignore
            If Action <> "File" Then
                _MainWindow.BlockDuplicateEventProcessing("Remove", "unknown")
            End If

            'Do the move
            Dim oMovedEmail As Microsoft.Office.Interop.Outlook.MailItem = mail.Move(targetFolder)

            'Get new Entry ID
            Dim MovedEntryID As String = oMovedEmail.EntryID

            'unless it was the result of an undo request, record the action
            If Action <> "Undo" Then
                LogAction(Action, FixedIndex, MovedEntryID, TargetStoreID, SourceEntryID, SourceStoreID)
            End If

            ReturnCode = MovedEntryID

            oMovedEmail = Nothing

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Outlook Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 ex.TargetSite.Name & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

            ReturnCode = ""
        Finally
            SetMousePointer(Cursors.Arrow)
        End Try

        Return ReturnCode

    End Function

    Private Sub LogAction(ByVal Action As String, ByVal FixedIndex As Integer, Optional ByVal MovedEntryID As String = "", Optional ByVal TargetStoreID As String = "", Optional ByVal SourceEntryID As String = "", Optional ByVal SourceStoreID As String = "")

        ' Record what happened (so that it can be undone later if necessary)

        If ActionLogIndex < 0 Then ActionLogIndex = 0

        Select Case Action
            Case Is = "File"
                ActionLog(ActionLogIndex, ActionLogSubIndex).ActionApplied = ActionType.File
            Case Is = "Delete"
                ActionLog(ActionLogIndex, ActionLogSubIndex).ActionApplied = ActionType.Delete
            Case Is = "Hide"
                ActionLog(ActionLogIndex, ActionLogSubIndex).ActionApplied = ActionType.Hide
        End Select
        ActionLog(ActionLogIndex, ActionLogSubIndex).FixedIndex = FixedIndex
        ActionLog(ActionLogIndex, ActionLogSubIndex).EmailID = MovedEntryID
        ActionLog(ActionLogIndex, ActionLogSubIndex).SourceStoreID = TargetStoreID
        ActionLog(ActionLogIndex, ActionLogSubIndex).TargetEntryID = SourceEntryID
        ActionLog(ActionLogIndex, ActionLogSubIndex).TargetStoreID = SourceStoreID

        ActionLogSubIndex += 1

        If Me.MenuUndo.IsEnabled Then
        Else
            MenuOptionEnabled("Undo", True)
        End If

    End Sub


    Private Function ConfirmActionMessage(ByVal strAction As String) As Boolean

        Select Case strAction

            Case Is = "Delete"
                If Not My.Settings.ConfirmDelete Then
                    Return True
                    Exit Function
                End If

            Case Is = "Exit"
                If Not My.Settings.ConfirmExit Then
                    Return True
                    Exit Function
                End If

            Case Is = "File"
                If Not My.Settings.ConfirmFile Then
                    Return True
                    Exit Function
                End If

            Case Is = "Hide"
                If Not My.Settings.ConfirmOpen Then
                    Return True
                    Exit Function
                End If

            Case Is = "Open"
                If Not My.Settings.ConfirmOpen Then
                    Return True
                    Exit Function
                End If

            Case Is = "Undo"
                If Not My.Settings.ConfirmOpen Then
                    Return True
                    Exit Function
                End If

        End Select

        Dim FunctionReturnCode As Boolean
        Dim ShowMessageBoxReturnCode As CustomDialog.CustomDialogResults
        Dim DefaultButton As CustomDialog.CustomDialogResults
        Dim Header As String = "FileFriendly - Confirm " & strAction
        Dim Instruction As String = ""
        Dim AdditionalDetail As String = ""

        Select Case strAction

            Case "File", "Delete", "Hide"

                Instruction = "Would you like to " & strAction.ToLower & " the "

                Dim NumberOfItems = ListView1.SelectedItems.Count
                Dim Numbers() As String = {"Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"}
                Select Case NumberOfItems
                    Case Is = 1
                        Instruction &= "selected e-mail?"
                    Case 2 To 10
                        Instruction &= Numbers(NumberOfItems).ToLower & " selected e-mails?"
                    Case Is > 1
                        Instruction &= NumberOfItems & " selected e-mails?"
                End Select

            Case Is = "Open"

                Instruction = "Would you like to open an e-mail?"
                AdditionalDetail = "If you have selected multiple e-mails, only the first one will be opened."

            Case Is = "Undo"

                Instruction = "Would you like to undo your last action?"

            Case Is = "Exit"

                Instruction = "Would you like to exit?"
                AdditionalDetail = "This prompt can be turned off in the Options Window."

        End Select

        DefaultButton = CustomDialog.CustomDialogResults.Yes

        ShowMessageBoxReturnCode = ShowMessageBox(Header,
            CustomDialog.CustomDialogIcons.Question,
            Instruction,
            "",
            AdditionalDetail,
            "",
            CustomDialog.CustomDialogIcons.None,
            CustomDialog.CustomDialogButtons.YesNo,
            DefaultButton)

        If ShowMessageBoxReturnCode = CustomDialog.CustomDialogResults.Yes Then
            FunctionReturnCode = True
        Else
            FunctionReturnCode = False
        End If

        Return FunctionReturnCode

    End Function

    Private Sub ShowFolderWindow()

        Try

            If gPickAFolderWindow Is Nothing Then
                gPickAFolderWindow = New PickAFolder
                gWhoIsInControl = WhoIsInControlType.Main

                If Me.WindowState = Windows.WindowState.Minimized Then
                    gMinimizedAtEarlyStartup = True
                End If

            Else
                gPickAFolderWindow.SafelyRefreshPickAFolderWindow()
            End If

            If gPickAFolderWindow.Visibility = Windows.Visibility.Hidden Then
                gPickAFolderWindow.Show()
            End If

        Catch ex As Exception
        End Try

    End Sub

    Private SelectedListViewItem As New ListViewRowClass

    Private Sub ActionRequestAgainstAllSelectedItems(ByVal Action As String, ByRef ListView1 As ListView)

        If ListView1.SelectedItems.Count = 0 Then Exit Sub

        Dim selectionSnapshot As SelectionSnapshot = CaptureSelectionSnapshot()

        Me.ForceCursor = True
        Me.Cursor = Cursors.Wait

        gSuppressUpdatesToDetailBox = True

        Try

            If ListView1.SelectedItems.Count > ActionLogMaxSubEntries Then

                Call ShowMessageBox("FileFriendly - Opps",
                         CustomDialog.CustomDialogIcons.Warning,
                         "Opps",
                         "You can only action " & ActionLogMaxSubEntries & " e-mails at a time.",
                         "You selected " & ListView1.SelectedItems.Count & " e-mails." & vbCrLf & "Please select fewer than " & ActionLogMaxSubEntries & " e-mails and redo your request.",
                        , , , CustomDialog.CustomDialogResults.OK)

                Exit Try
            End If

            'Prepare to action all selected entries
            Dim SelectedEntries(ListView1.SelectedItems.Count) As Integer

            Dim Count As Integer = 0
            For Each SelectedItem In ListView1.SelectedItems
                SelectedEntries(Count) = SelectedItem.Index
                Count += 1
            Next

            If Count > 0 Then
                Array.Sort(SelectedEntries)
                Array.Reverse(SelectedEntries)
                ActionRequest_Worker(Action, SelectedEntries, Count, ListView1)
            End If

        Catch ex As Exception

        End Try

        gSuppressUpdatesToDetailBox = False

        RestoreSelection(selectionSnapshot, SelectionRestoreReason.UserDelete, True)

        Me.Cursor = Cursors.None

    End Sub

    Private Function GetDeleteFolderIndexForStore(ByVal storeId As String) As Integer

        If Not String.IsNullOrEmpty(storeId) Then
            Dim info As StoreDeleteFolderInfo = Nothing
            If gStoreDeleteFolders.TryGetValue(storeId, info) Then
                If info.FolderIndex >= 0 AndAlso info.FolderIndex < gFolderTable.Length Then
                    Return info.FolderIndex
                End If
            End If
        End If

        ' Fallback: use global deleted folder index if available
        If gDeletedFolderIndex >= 0 AndAlso gDeletedFolderIndex < gFolderTable.Length Then
            Return gDeletedFolderIndex
        End If

        ' Absolute fallback: use original folder (no move) to avoid crashing
        Return 0

    End Function
    Private Sub ActionRequest_Worker(ByVal Action As String, ByRef SelectedEntries() As Integer, ByVal Count As Integer, ByRef ListView1 As ListView)

        Static Dim TooManyActionsMessageDisplayed As Boolean = False

        Try

            'action requests
            Dim IndexToAction As Integer

            For z As Integer = 0 To Count - 1

                IndexToAction = SelectedEntries(z)

                Select Case Action

                    Case "File", "Delete"

                        If gPickFromContextMenuOverride > 0 Then

                            If ListView1.Items(IndexToAction).OriginalFolder = gPickFromContextMenuOverride Then

                                LogAction("Hide", ListView1.Items(IndexToAction).FixedIndex)

                            Else

                                Dim EntryIDFromListView As String = ListView1.Items(IndexToAction).OutlookEntryID

                                Dim x As Integer

                                x = ListView1.Items(IndexToAction).OriginalFolder
                                Dim SourceEntryID As String = gFolderTable(x).EntryID
                                Dim SourceStoreID As String = gFolderTable(x).StoreID

                                If Action = "File" Then
                                    x = gPickFromContextMenuOverride
                                Else
                                    ' Delete: choose a per‑store Deleted/Trash folder
                                    Dim sourceStoreIdForDelete As String = SourceStoreID
                                    x = GetDeleteFolderIndexForStore(sourceStoreIdForDelete)
                                End If

                                Dim TargetEntryID As String = gFolderTable(x).EntryID
                                Dim TargetStoreID As String = gFolderTable(x).StoreID

                                'File the message
                                Dim newId As String = FileMessage(Action,
                                                  ListView1.Items(IndexToAction).FixedIndex,
                                                  EntryIDFromListView,
                                                  SourceEntryID,
                                                  SourceStoreID,
                                                  TargetEntryID,
                                                  TargetStoreID)

                                ' If the move/delete failed, do NOT remove the row
                                If String.IsNullOrEmpty(newId) Then
                                    ' Abort processing remaining items; user already saw an error
                                    Exit For
                                End If

                            End If
                        Else

                            If (ListView1.Items(IndexToAction).OriginalFolder = ListView1.SelectedItem.RecommendedFolderFinal) Then

                                LogAction("Hide", ListView1.Items(IndexToAction).FixedIndex)

                            Else

                                Dim EmailID As String = ListView1.Items(IndexToAction).OutlookEntryID

                                Dim x As Integer

                                x = ListView1.Items(IndexToAction).OriginalFolder
                                Dim SourceEntryID As String = gFolderTable(x).EntryID
                                Dim SourceStoreID As String = gFolderTable(x).StoreID

                                If Action = "File" Then
                                    x = ListView1.Items(IndexToAction).RecommendedFolderFinal
                                Else
                                    ' Delete: choose a per‑store Deleted/Trash folder
                                    Dim sourceStoreIdForDelete As String = SourceStoreID
                                    x = GetDeleteFolderIndexForStore(sourceStoreIdForDelete)
                                End If

                                Dim TargetEntryID As String = gFolderTable(x).EntryID
                                Dim TargetStoreID As String = gFolderTable(x).StoreID

                                'File the message
                                Dim newId As String = FileMessage(Action,
                                                  ListView1.Items(IndexToAction).FixedIndex,
                                                  EmailID,
                                                  SourceEntryID,
                                                  SourceStoreID,
                                                  TargetEntryID,
                                                  TargetStoreID)

                                ' If the move/delete failed, do NOT remove the row
                                If String.IsNullOrEmpty(newId) Then
                                    Exit For
                                End If

                            End If

                        End If

                        ' Only remove the row from the grid if the action succeeded
                        RemoveAnEntry(IndexToAction)

                    Case "Hide"

                        LogAction(Action, ListView1.Items(IndexToAction).FixedIndex)
                        RemoveAnEntry(IndexToAction)

                End Select

            Next z

            ActionLogSubIndex = 0
            ActionLogIndex += 1

            If ActionLogIndex > ActionLogMaxEntries Then

                If TooManyActionsMessageDisplayed Then
                Else
                    TooManyActionsMessageDisplayed = True

                    Call ShowMessageBox("FileFriendly - Opps",
                        "It looks like your on a roll so you can keep on going, but you will only be able to undo your most recent " & ActionLogMaxEntries & " actions from now on.",
                      , , , CustomDialog.CustomDialogResults.OK)

                End If

                'clear action left over action log entries
                For i As Integer = 1 To ActionLogMaxEntries
                    For ii As Integer = 0 To ActionLogMaxSubEntries
                        If ActionLog(i - 1, ii).EmailID = Nothing Then
                            If ActionLog(i, ii).EmailID = Nothing Then
                                Exit For
                            End If
                        End If
                        ActionLog(i - 1, ii) = ActionLog(i, ii)
                    Next
                Next
                For ii As Integer = 0 To ActionLogMaxSubEntries
                    ActionLog(ActionLogMaxEntries, ii).EmailID = Nothing
                    ActionLog(ActionLogMaxEntries, ii).SourceStoreID = Nothing
                    ActionLog(ActionLogMaxEntries, ii).TargetEntryID = Nothing
                    ActionLog(ActionLogMaxEntries, ii).TargetStoreID = Nothing
                Next

                ActionLogIndex -= 1

            End If

            ReindexListView(ListView1)

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly",
                     CustomDialog.CustomDialogIcons.Stop,
                     "Unexpected Error!",
                     "FileFriendly has encountered an unexpected error.",
                     ex.TargetSite.Name & " - " & ex.ToString,
                     "",
                     CustomDialog.CustomDialogIcons.None,
                     CustomDialog.CustomDialogButtons.OK,
                     CustomDialog.CustomDialogResults.OK)

        End Try

    End Sub

    Private Sub ReindexListView(ByRef lv As ListView)

        For x As Integer = 0 To lv.Items.Count - 1
            lv.Items(x).index = x
        Next

    End Sub

    Private Sub RemoveAnEntry(ByVal IndexToAction As Integer)

        ' Remove an entry from the:
        ' 1. List view
        ' 2. final recommendation table
        ' 3. email table

        Dim entryIDToBeRemoved As String = ListView1.Items(IndexToAction).OutlookEntryID

        ' Remove from the list view
        ListView1.Items.RemoveAt(IndexToAction)

        ' remove from the final recommendation table; resize the table when done
        Dim entryFound As Boolean

        entryFound = False
        For i = 1 To gFinalRecommendationTable.Length - 1

            If gFinalRecommendationTable(i - 1).OutlookEntryID = entryIDToBeRemoved Then
                entryFound = True
                Continue For
            End If

            If entryFound Then
                gFinalRecommendationTable(i - 1) = gFinalRecommendationTable(i)
            End If

        Next

        If entryFound Then
            If gFinalRecommendationTable.Length > 1 Then
                ReDim Preserve gFinalRecommendationTable(gFinalRecommendationTable.Length - 2)
            Else
                ReDim gFinalRecommendationTable(0)
            End If
        End If

        ' Remove the entry for the email table; resize the table when done

        entryFound = False
        For i = 1 To gEmailTable.Length - 1

            If gEmailTable(i).sOutlookEntryID = entryIDToBeRemoved Then
                entryFound = True
                gEmailTableIndex -= 1
                Continue For
            End If

            If entryFound Then
                gEmailTable(i - 1) = gEmailTable(i)
            End If

        Next

        If entryFound Then
            If gEmailTable.Length > 1 Then
                ReDim Preserve gEmailTable(gEmailTable.Length - 2)
            Else
                ReDim gEmailTable(0)
            End If
        End If


    End Sub

#Region "Ensure Outlook is running if needed"

    ' Non-blocking status notification used when starting Outlook
    Private Sub ShowOutlookStartingMessage()
        Try
            Me.Dispatcher.BeginInvoke(
                New SetFolderNameTextCallback(AddressOf SetFoldersNameText),
                New Object() {"Starting Outlook – please wait…"})
        Catch
        End Try
    End Sub

    Private Sub ClearOutlookStartingMessage()
        Try
            ' Only clear if we are still showing the "Starting Outlook" text
            'If TypeOf Me.lblMainMessageLine.Content Is String Then
            '    Dim current As String = CStr(Me.lblMainMessageLine.Content)
            '    If current.Contains("Starting Outlook") Then
            'Me.Dispatcher.BeginInvoke(
            '            New SetFolderNameTextCallback(AddressOf SetFoldersNameText),
            '            New Object() {"0 e-mails"})

            Try
                Me.Dispatcher.BeginInvoke(
                New SetFolderNameTextCallback(AddressOf SetFoldersNameText),
                New Object() {"Outlook started"})
            Catch
            End Try

            '    End If
            'End If
        Catch
        End Try
    End Sub

    Private Function GetCurrentOutlookProcessId() As Integer
        Try
            Dim currentSessionId As Integer = Process.GetCurrentProcess().SessionId
            For Each p As Process In Process.GetProcessesByName("OUTLOOK")
                Try
                    If p.SessionId = currentSessionId AndAlso Not p.HasExited Then
                        If p.MainWindowHandle <> IntPtr.Zero Then
                            Return p.Id
                        End If
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
        Return -1
    End Function

    Private Function IsOutlookProcessRunning() As Boolean
        Return GetCurrentOutlookProcessId() > 0
    End Function

    Private Function EnsureOutlookIsRunning() As Boolean

        If Not Me.Dispatcher.CheckAccess() Then
            Return CBool(Me.Dispatcher.Invoke(New Func(Of Boolean)(AddressOf EnsureOutlookIsRunning)))
        End If

        Dim originalCursor As System.Windows.Input.Cursor = Nothing
        Try
            originalCursor = Me.Cursor
        Catch
        End Try

        SetMousePointer(Cursors.Wait)

        Try
            Dim repairOnly As Boolean = False

            If oApp IsNot Nothing AndAlso oNS IsNot Nothing Then

                If Not System.Runtime.InteropServices.Marshal.IsComObject(oNS) Then
                    repairOnly = True
                    oNS = Nothing
                    oApp = Nothing
                Else
                    Try
                        Dim dummy As Integer = oNS.Folders.Count
                        Return True
                    Catch invalidEx As System.Runtime.InteropServices.InvalidComObjectException
                        repairOnly = True
                        oNS = Nothing
                        oApp = Nothing
                    Catch comEx As System.Runtime.InteropServices.COMException
                        Const RPC_E_SERVER_UNAVAILABLE As Integer = &H800706BA
                        Const RPC_E_DISCONNECTED As Integer = &H800706BE

                        If comEx.HResult = RPC_E_SERVER_UNAVAILABLE OrElse comEx.HResult = RPC_E_DISCONNECTED Then
                            repairOnly = True
                            oNS = Nothing
                            oApp = Nothing
                        Else
                            Throw
                        End If
                    End Try
                End If
            End If

            If Not repairOnly AndAlso Not IsOutlookProcessRunning() Then

                Dim header As String = "FileFriendly - Start Outlook?"
                Dim instruction As String = vbCrLf &
                    "Microsoft Outlook is not running." & vbCrLf & vbCrLf &
                    "FileFriendly needs Outlook to be running to help file your e-mails." & vbCrLf & vbCrLf &
                    "Would you like FileFriendly to automatically start Outlook for you now?"
                Dim detail As String =
                    "If you choose 'Yes', FileFriendly will automatically start Outlook." & vbCrLf & vbCrLf &
                    "If you choose 'No', FileFriendly will close." & vbCrLf & "Later, if you wish, you can manually start Outlook and then FileFriendly."


                If My.Settings.SoundAlert Then Beep()

                Dim response As CustomDialog.CustomDialogResults =
                    ShowMessageBox(header,
                                   CustomDialog.CustomDialogIcons.Question,
                                   instruction,
                                   "",
                                   detail,
                                   "",
                                   CustomDialog.CustomDialogIcons.None,
                                   CustomDialog.CustomDialogButtons.YesNo,
                                   CustomDialog.CustomDialogResults.Yes)

                If response <> CustomDialog.CustomDialogResults.Yes Then
                    ShowMessageBox("FileFriendly",
                                   CustomDialog.CustomDialogIcons.Information,
                                   "FileFriendly will close when you click 'OK'.",
                                   "",
                                   "To run FileFriendly ideally Outlook should be already be running.",
                                   "",
                                   CustomDialog.CustomDialogIcons.None,
                                   CustomDialog.CustomDialogButtons.OK,
                                   CustomDialog.CustomDialogResults.OK)

                    Application.Current.Shutdown() ' exit the program now
                    Return False
                End If

                ShowOutlookStartingMessage()

                Try
                    Dim startInfo As New ProcessStartInfo("outlook.exe")
                    Process.Start(startInfo)

                Catch exStart As Exception

                    ClearOutlookStartingMessage()

                    Call ShowMessageBox("FileFriendly - Outlook start fail",
                        CustomDialog.CustomDialogIcons.Stop,
                        "Unexpected Error!",
                        "The requested action cannot be completed until Outlook is available.",
                        exStart.TargetSite.Name & " - " & exStart.ToString,
                        "",
                        CustomDialog.CustomDialogIcons.None,
                        CustomDialog.CustomDialogButtons.OK,
                        CustomDialog.CustomDialogResults.OK)

                    Return False

                End Try

                Dim startedOk As Boolean = False
                Dim deadline As Date = Date.UtcNow.AddSeconds(60)

                Do
                    System.Threading.Thread.Sleep(500)

                    Try
                        Dim currentSessionId As Integer = Process.GetCurrentProcess().SessionId
                        For Each p As Process In Process.GetProcessesByName("OUTLOOK")
                            Try
                                If p.SessionId = currentSessionId AndAlso Not p.HasExited Then
                                    If p.MainWindowHandle <> IntPtr.Zero Then
                                        startedOk = True
                                        Exit For
                                    End If
                                End If
                            Catch
                            End Try
                        Next
                    Catch
                    End Try

                    If startedOk Then Exit Do
                Loop While Date.UtcNow < deadline

                ScheduleRefreshGrid()

            End If

            Try
                oApp = CType(CreateObject("Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
            Catch exCreate As Exception

                Call ShowMessageBox("FileFriendly - Outlook error",
                    CustomDialog.CustomDialogIcons.Stop,
                    "Unexpected Error!",
                    "FileFriendly has encountered an unexpected error." & vbCrLf & "CreateObject('Outlook.Application')",
                    exCreate.TargetSite.Name & " - " & exCreate.ToString,
                    "",
                    CustomDialog.CustomDialogIcons.None,
                    CustomDialog.CustomDialogButtons.OK,
                    CustomDialog.CustomDialogResults.OK)

                oApp = Nothing
                oNS = Nothing
                Return False
            End Try

            Try
                oNS = oApp.GetNamespace("MAPI")
                Dim dummy As Integer = oNS.Folders.Count
            Catch exNs As System.Exception

                Call ShowMessageBox("FileFriendly - Outlook error",
                    CustomDialog.CustomDialogIcons.Stop,
                    "Unexpected Error!",
                    "FileFriendly has encountered an unexpected error." & vbCrLf & "oApp.GetNamespace('MAPI').",
                    exNs.TargetSite.Name & " - " & exNs.ToString,
                    "",
                    CustomDialog.CustomDialogIcons.None,
                    CustomDialog.CustomDialogButtons.OK,
                    CustomDialog.CustomDialogResults.OK)
                oNS = Nothing
                oApp = Nothing
                Return False
            End Try

            Return True

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly - Outlook error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 ex.TargetSite.Name & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

            oNS = Nothing
            oApp = Nothing
            Return False

        Finally
            If originalCursor IsNot Nothing Then
                SetMousePointer(originalCursor)
            Else
                SetMousePointer(Cursors.Arrow)
            End If
        End Try

    End Function
    Private Sub OnOutlookQuit()

        Try

            Dim previousProcessId As Integer = GetCurrentOutlookProcessId()

            If gOutlookEventHandler IsNot Nothing Then
                gOutlookEventHandler.Dispose()
                gOutlookEventHandler = Nothing
            End If
            oNS = Nothing
            oApp = Nothing

            Dim restartThread As New Thread(Sub()
                                                Try
                                                    Dim quitDeadline As Date = Date.UtcNow.AddSeconds(30)
                                                    If previousProcessId > 0 Then
                                                        Do
                                                            Try
                                                                Dim p As Process = Process.GetProcessById(previousProcessId)
                                                                If p.HasExited Then Exit Do
                                                            Catch
                                                                Exit Do
                                                            End Try
                                                            If Date.UtcNow >= quitDeadline Then Exit Do
                                                            Thread.Sleep(500)
                                                        Loop
                                                    End If

                                                    Me.Dispatcher.Invoke(Sub()
                                                                             If EnsureOutlookIsRunning() Then
                                                                                 InitializeMonitoringOfOutlookEvents()
                                                                             End If
                                                                         End Sub)
                                                Catch
                                                End Try
                                            End Sub)
            restartThread.IsBackground = True
            restartThread.SetApartmentState(System.Threading.ApartmentState.STA)
            restartThread.Start()

        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try

    End Sub

#End Region

    Private Sub imgClose_MouseLeftButtonDown1(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles imgClose.MouseLeftButtonDown

        If My.Settings.ConfirmExit Then
            If ConfirmActionMessage("Exit") Then ShutDown()
        Else
            ShutDown()
        End If

    End Sub

    Private Sub imgMinimize_MouseLeftButtonDown1(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles imgMinimize.MouseLeftButtonDown

        If gMinimizeMaximizeAllowed Then
            Me.ShowInTaskbar = True
            Me.WindowState = Windows.WindowState.Minimized
            If gPickAFolderWindow IsNot Nothing Then
                gPickAFolderWindow.WindowState = Windows.WindowState.Minimized
            End If
        Else
            Beep()
        End If

    End Sub
    Private Sub imgMaximize_MouseLeftButtonDown1(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles imgMaximize.MouseLeftButtonDown

        Static Dim OldTopLocation As Double = 0

        Try

            If gMinimizeMaximizeAllowed Then

                Dim imageUri As Uri

                If Me.WindowState = Windows.WindowState.Maximized Then

                    Me.Top = OldTopLocation
                    gMainWindowIsMaximized = False
                    gPickAFolderWindow.SafelyMakePickAFolderWindowTopMost()
                    Me.WindowState = Windows.WindowState.Normal
                    Me.ShowInTaskbar = True
                    imgMaximize.ToolTip = "Maximize"
                    imageUri = New Uri("/filefriendly;component/Resources/maximize.gif", UriKind.Relative)

                Else

                    OldTopLocation = Me.Top
                    gPickAFolderWindowWasDocedWhenMainWindowWasMaximimized = gWindowDocked
                    gMainWindowIsMaximized = True
                    gPickAFolderWindow.SafelyMakePickAFolderWindowTopMost()
                    Me.WindowState = Windows.WindowState.Maximized
                    imgMaximize.ToolTip = "Restore"
                    imageUri = New Uri("/filefriendly;component/Resources/restore.gif", UriKind.Relative)

                End If

                Dim BitmapSource As BitmapSource = New BitmapImage(imageUri)
                imgMaximize.Source = BitmapSource

            Else

                Beep()

            End If

        Catch ex As Exception
        End Try

    End Sub

    Private Sub MainWindow_SizeChanged(ByVal sender As Object, ByVal e As System.Windows.SizeChangedEventArgs) Handles Me.SizeChanged

        gmwHeight = Me.ActualHeight
        gmwWidth = Me.ActualWidth

        ' Recalculate list view columns when main window size changes
        If Me.ListView1 IsNot Nothing AndAlso Me.ListView1.ActualWidth > 0 Then
            RecalculateListViewColumnWidths()
        End If

    End Sub

    Private Sub UpdateMailboxColumnVisibility()

        If Not Dispatcher.CheckAccess() Then
            Dispatcher.BeginInvoke(New Action(AddressOf UpdateMailboxColumnVisibility))
            Return
        End If

        Dim gridView As GridView = TryCast(ListView1.View, GridView)
        If gridView Is Nothing OrElse gridView.Columns.Count = 0 Then
            Return
        End If

        Const mailboxColumnIndex As Integer = 1
        If gridView.Columns.Count <= mailboxColumnIndex Then
            Return
        End If

        Dim mailboxColumn As GridViewColumn = gridView.Columns(mailboxColumnIndex)

        'If _mailboxCount <= 1 Then

        If _TotalMailBoxes <= 1 Then
            ' Single mailbox: hide column
            mailboxColumn.Width = 0
            mailboxColumn.Header = String.Empty
        Else
            ' Multiple mailboxes: ensure header text is present
            If String.IsNullOrEmpty(TryCast(mailboxColumn.Header, String)) Then
                mailboxColumn.Header = "Mailbox"
            End If
        End If

    End Sub

    Private Sub RecalculateListViewColumnWidths()

        Dim gridView As GridView = TryCast(ListView1.View, GridView)
        If gridView Is Nothing OrElse gridView.Columns.Count = 0 Then
            Return
        End If

        If ListView1.ActualWidth <= 0 Then
            Return
        End If

        ' Column indices:
        ' 0 = chain/indicator 1 (fixed)
        ' 1 = Mailbox (auto/hide)
        ' 2 = Subject (variable)
        ' 3 = From (variable)
        ' 4 = To (variable)
        ' 5 = Date (fixed)
        ' 6 = Time (fixed)
        ' 7 = chain/indicator 2 (fixed)

        If gridView.Columns.Count < 8 Then
            ' Layout assumptions not met – do nothing
            Return
        End If

        Dim chain1Column As GridViewColumn = gridView.Columns(0)
        Dim mailboxColumn As GridViewColumn = gridView.Columns(1)
        Dim subjectColumn As GridViewColumn = gridView.Columns(2)
        Dim fromColumn As GridViewColumn = gridView.Columns(3)
        Dim toColumn As GridViewColumn = gridView.Columns(4)
        Dim dateColumn As GridViewColumn = gridView.Columns(5)
        Dim timeColumn As GridViewColumn = gridView.Columns(6)
        Dim chain2Column As GridViewColumn = gridView.Columns(7)

        ' DPI for FormattedText measurement (fallback for older frameworks)
        Dim pixelsPerDip As Double = 1.0

        Dim baseTypeface As New Typeface(Me.FontFamily, Me.FontStyle, Me.FontWeight, Me.FontStretch)
        Dim boldTypeface As New Typeface(Me.FontFamily, Me.FontStyle, FontWeights.Bold, Me.FontStretch)

        Dim headerPadding As Double = 16  ' approximate header padding
        Dim cellPadding As Double = 14    ' approximate cell padding

        Dim measureText As Func(Of String, Typeface, Double) =
            Function(text As String, typeface As Typeface) As Double
                If String.IsNullOrEmpty(text) Then
                    Return 0
                End If
                Dim ft As New FormattedText(
                    text,
                    System.Globalization.CultureInfo.CurrentCulture,
                    FlowDirection.LeftToRight,
                    typeface,
                    Me.FontSize,
                    System.Windows.Media.Brushes.Black,
                    New NumberSubstitution(),
                    TextFormattingMode.Display,
                    pixelsPerDip)
                Return ft.WidthIncludingTrailingWhitespace
            End Function

        ' ---------- 1. Fixed-width columns ----------

        Dim chain1Header As String = TryCast(chain1Column.Header, String)
        Dim chain2Header As String = TryCast(chain2Column.Header, String)
        Dim dateHeader As String = TryCast(dateColumn.Header, String)
        'Dim timeHeader As String = TryCast(timeColumn.Header, String)

        Dim minFixedWidth As Double = 18

        Dim chain1Width As Double = Math.Max(minFixedWidth, measureText(chain1Header, baseTypeface) + headerPadding)
        Dim chain2Width As Double = Math.Max(minFixedWidth, measureText(chain2Header, baseTypeface) + headerPadding)

        Dim dateWidth As Double = Math.Max(80, measureText(If(dateHeader, "Date"), baseTypeface) + headerPadding)

        ' Time width based on the widest possible time string for the current format,
        ' measured in both normal and bold, then taking the larger.
        Dim widestTimeSample = New Date(2000, 12, 31, 23, 59, 59)
        Dim widestTimeText As String

        Try
            widestTimeText = Format(widestTimeSample, gPreferredTimeFormat)
        Catch
            widestTimeText = "23:59:59 PM"
        End Try

        Dim normalTimeWidth As Double = measureText(widestTimeText, baseTypeface)
        Dim boldTimeWidth As Double = measureText(widestTimeText, boldTypeface)

        Dim measuredTimeWidth As Double = Math.Max(normalTimeWidth, boldTimeWidth)

        ' Add padding and enforce a small absolute minimum
        Dim timeWidth As Double = Math.Max(70, measuredTimeWidth + cellPadding)

        chain1Column.Width = chain1Width
        chain2Column.Width = chain2Width
        dateColumn.Width = dateWidth
        timeColumn.Width = timeWidth

        ' ---------- 2. Mailbox column (content-based, or hidden) ----------

        Dim mailboxVisible As Boolean = (_TotalMailBoxes > 1)
        Dim mailboxWidth As Double = 0

        If mailboxVisible Then
            If String.IsNullOrEmpty(TryCast(mailboxColumn.Header, String)) Then
                mailboxColumn.Header = "Mailbox"
            End If

            Dim maxMailboxTextWidth As Double = 0

            For Each obj In ListView1.Items
                Dim row As ListViewRowClass = TryCast(obj, ListViewRowClass)
                If row Is Nothing OrElse String.IsNullOrEmpty(row.MailBoxName) Then
                    Continue For
                End If
                Dim w As Double = measureText(row.MailBoxName, baseTypeface)
                If w > maxMailboxTextWidth Then
                    maxMailboxTextWidth = w
                End If
            Next

            Dim mailboxHeaderText As String = TryCast(mailboxColumn.Header, String)
            Dim headerWidth As Double = measureText(mailboxHeaderText, baseTypeface)

            Dim contentBased As Double = Math.Max(maxMailboxTextWidth, headerWidth)

            mailboxWidth = contentBased + cellPadding
            mailboxWidth = Math.Max(mailboxWidth, 60)

            mailboxColumn.Width = mailboxWidth
        Else
            mailboxColumn.Width = 0
            mailboxColumn.Header = String.Empty
        End If

        ' ---------- 3. Allocate remaining width to Subject / From / To ----------

        Dim totalAvailable As Double = ListView1.ActualWidth
        Dim layoutMargin As Double = 30
        totalAvailable = Math.Max(0, totalAvailable - layoutMargin)

        Dim fixedTotal As Double = chain1Width + chain2Width + dateWidth + timeWidth + mailboxWidth
        Dim remaining As Double = totalAvailable - fixedTotal

        If remaining <= 0 Then
            subjectColumn.Width = 80
            fromColumn.Width = 60
            toColumn.Width = 60
            Return
        End If

        Dim totalWeight As Double = 4 + 2 + 2
        Dim subjectShare As Double = 4 / totalWeight
        Dim fromShare As Double = 2 / totalWeight
        Dim toShare As Double = 2 / totalWeight

        Dim subjectWidth As Double = remaining * subjectShare
        Dim fromWidth As Double = remaining * fromShare
        Dim toWidth As Double = remaining * toShare

        subjectWidth = Math.Max(120, subjectWidth)
        fromWidth = Math.Max(80, fromWidth)
        toWidth = Math.Max(80, toWidth)

        Dim currentVariableTotal As Double = subjectWidth + fromWidth + toWidth
        If currentVariableTotal > remaining Then
            Dim scale As Double = remaining / currentVariableTotal
            subjectWidth *= scale
            fromWidth *= scale
            toWidth *= scale
        End If

        subjectColumn.Width = subjectWidth
        fromColumn.Width = fromWidth
        toColumn.Width = toWidth

    End Sub

    Private Sub MainWindow_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged, Me.SizeChanged

        Static Dim LastState As WindowState = Windows.WindowState.Normal

        ' when just restoring to a normal state or when going into a minimized or maximized state 
        ' don't do anything because top and left values are not usable

        If (LastState = Windows.WindowState.Minimized) Or (LastState = Windows.WindowState.Maximized) Then
            ' if last state was minimized or maximized 
            ' then we are just restoring now
            ' so don't do anything
        Else
            If Me.WindowState = Windows.WindowState.Normal Then
                gmwTop = MainWindow.Top
                gmwLeft = MainWindow.Left
                If gPickAFolderWindow IsNot Nothing Then
                    If gWindowDocked Then gPickAFolderWindow.SafelyMovePickAFolderWindow()
                End If
            End If

        End If

        LastState = Me.WindowState

    End Sub

    Private Sub MainWindow_StateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.StateChanged

        On Error Resume Next

        'restore to last position
        If Me.WindowState = Windows.WindowState.Normal Then
            Me.Top = gmwTop
            Me.Left = gmwLeft
            Me.Width -= 1 : Me.Width += 1 ' bump the window so the columns align
        End If

        If gPickAFolderWindow IsNot Nothing Then
            If Me.WindowState = Windows.WindowState.Minimized Then
                gPickAFolderWindow.SafelyHidePickAFolderWindow()
            Else
                gPickAFolderWindow.SafelyShowPickAFolderWindow()
                gPickAFolderWindow.SafelyMakePickAFolderWindowTopMost()
            End If
        End If

    End Sub

    Enum ListSortDirection
        Ascending = 1
        Descending = 2
    End Enum
    Private _lastDirection As ListSortDirection = ListSortDirection.Descending
    Private _lastheaderClicked As GridViewColumnHeader
    Private _hasUserSorted As Boolean = False
    Private gCurrentSortDirection As ListSortDirection = ListSortDirection.Ascending

    Private Sub ApplyCurrentSortOrderToFinalTable()

        If gFinalRecommendationTable Is Nothing OrElse gFinalRecommendationTable.Length = 0 Then Return

        Dim column As String
        Dim direction As FinalRecommendationTableSorter.MySortOrder

        If _hasUserSorted Then
            column = gCurrentSortOrder
            direction = If(gCurrentSortDirection = ListSortDirection.Descending,
                           FinalRecommendationTableSorter.MySortOrder.Descending,
                           FinalRecommendationTableSorter.MySortOrder.Ascending)
        Else
            column = "Mailbox"
            direction = FinalRecommendationTableSorter.MySortOrder.Ascending
            gCurrentSortDirection = ListSortDirection.Ascending
        End If

        Dim sorter As New FinalRecommendationTableSorter With {
          .PrimaryColumnToSort = column,
          .SortOrder = direction
        }

        Array.Sort(gFinalRecommendationTable, sorter)

        If _hasUserSorted Then
            gCurrentSortOrder = column
        End If

        UpdateSortHeaderGlyph()

    End Sub

    Private Sub UpdateSortHeaderGlyph()

        If Not Dispatcher.CheckAccess() Then
            Dispatcher.BeginInvoke(New Action(AddressOf UpdateSortHeaderGlyph))
            Return
        End If

        If Not _hasUserSorted Then Return

        Dim gv As GridView = TryCast(ListView1.View, GridView)
        If gv Is Nothing Then Return

        Dim targetHeader As String = gCurrentSortOrder

        For Each col As GridViewColumn In gv.Columns
            Dim headerText As String = TryCast(col.Header, String)
            If String.IsNullOrWhiteSpace(headerText) Then Continue For

            Dim isMatch As Boolean = String.Equals(headerText, targetHeader, StringComparison.OrdinalIgnoreCase)
            If Not isMatch AndAlso String.Equals(targetHeader, "MailBoxName", StringComparison.OrdinalIgnoreCase) Then
                isMatch = String.Equals(headerText, "Mailbox", StringComparison.OrdinalIgnoreCase)
            End If

            If isMatch Then
                col.HeaderTemplate = If(gCurrentSortDirection = ListSortDirection.Descending,
                                        TryCast(Resources("HeaderTemplateArrowDown"), DataTemplate),
                                        TryCast(Resources("HeaderTemplateArrowUp"), DataTemplate))
            Else
                col.HeaderTemplate = Nothing
            End If
        Next

    End Sub

    Private Sub ListViewColumnHeaderClickedHandler(ByVal sender As Object, ByVal e As RoutedEventArgs)

        SetMousePointer(Cursors.Wait)

        Try

            StorePendingSelection(SelectionRestoreReason.Sort)

            Dim headerClicked As GridViewColumnHeader = TryCast(e.OriginalSource, GridViewColumnHeader)
            Dim direction As ListSortDirection

            If headerClicked IsNot Nothing Then

                If headerClicked.Role <> GridViewColumnHeaderRole.Padding Then

                    Dim header As String = TryCast(headerClicked.Column.Header, String)

                    If header.Trim.Length = 0 Then ' the header over the chain indicator was clicked
                        Exit Try
                    End If

                    If headerClicked IsNot _lastheaderClicked Then
                        direction = ListSortDirection.Ascending
                    Else
                        If _lastDirection = ListSortDirection.Ascending Then
                            direction = ListSortDirection.Descending
                        Else
                            direction = ListSortDirection.Ascending
                        End If
                    End If

                    Dim lFinalRecommendationTableSorter As New FinalRecommendationTableSorter With {
                        .PrimaryColumnToSort = header,
                        .SortOrder = direction
                    }
                    Array.Sort(gFinalRecommendationTable, lFinalRecommendationTableSorter)
                    lFinalRecommendationTableSorter = Nothing

                    _hasUserSorted = True
                    gCurrentSortDirection = direction
                    gCurrentSortOrder = header
                    SetListViewItem(gFinalRecommendationTable)

                    If direction = ListSortDirection.Ascending Then
                        headerClicked.Column.HeaderTemplate = TryCast(Resources("HeaderTemplateArrowUp"), DataTemplate)
                    Else
                        headerClicked.Column.HeaderTemplate = TryCast(Resources("HeaderTemplateArrowDown"), DataTemplate)
                    End If

                    ' Remove arrow from previously sorted header
                    If _lastheaderClicked IsNot Nothing AndAlso _lastheaderClicked IsNot headerClicked Then
                        _lastheaderClicked.Column.HeaderTemplate = Nothing
                    End If

                    _lastheaderClicked = headerClicked
                    _lastDirection = direction

                    UpdateSortHeaderGlyph()

                End If

            End If

            ApplyFilter()

        Catch ex As Exception

            Call ShowMessageBox("FileFriendly",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try

        SetMousePointer(Cursors.Arrow)

    End Sub

#Region "Real-Time Email Monitoring"

    Private gOutlookEventHandler As OutlookEventHandler
    Private ReadOnly gEventHandlerLock As New Object()

    Public Sub InitializeMonitoringOfOutlookEvents()

        Try

            If gFolderTable Is Nothing OrElse gFolderTable.Length = 0 Then
                Exit Sub
            End If

            If oNS Is Nothing Then
                If Not EnsureOutlookIsRunning() Then Exit Sub
                If oNS Is Nothing Then Exit Sub
            End If

            If GetCurrentOutlookProcessId() <= 0 Then Exit Sub

            Dim monitoringThread As New Thread(Sub()
                                                   Try
                                                       Dim ns As Microsoft.Office.Interop.Outlook.NameSpace = oNS

                                                       SyncLock gEventHandlerLock
                                                           If gOutlookEventHandler IsNot Nothing Then Return
                                                           ClearMonitoringOfOutlookEvents()
                                                           If ns Is Nothing Then Return
                                                           gOutlookEventHandler = New OutlookEventHandler(Me, ns)
                                                       End SyncLock
                                                   Catch ex As Exception
                                                   End Try
                                               End Sub) With {
            .IsBackground = True
            }
            monitoringThread.SetApartmentState(System.Threading.ApartmentState.STA)
            monitoringThread.Start()
            Thread.Sleep(500)

        Catch ex As Exception
        End Try

    End Sub

    Private gRefreshGridScheduled As Boolean = False
    Private ReadOnly gRefreshGridLock As New Object()

    ' Thread-safe helpers for gSuppressEventForEntryIds / _suspensionTimers.
    Private Function GetSuppressedEventCount() As Integer
        SyncLock gSuppressEventLock
            Return gSuppressEventForEntryIds.Count
        End SyncLock
    End Function

    Private Function IsEntryIdSuppressed(ByVal entryId As String) As Boolean
        If String.IsNullOrEmpty(entryId) Then Return False
        SyncLock gSuppressEventLock
            Return gSuppressEventForEntryIds.Contains(entryId)
        End SyncLock
    End Function

    Private Sub ScheduleRefreshGrid(Optional ByVal selectionReason As SelectionRestoreReason = SelectionRestoreReason.Refresh)

        Dim shouldSchedule As Boolean = False
        Dim captureSelection As Boolean = False

        SyncLock gRefreshGridLock
            If Not gRefreshGridScheduled Then
                gRefreshGridScheduled = True
                shouldSchedule = True
                captureSelection = True
            End If
        End SyncLock

        If Not shouldSchedule Then Return

        BeginRefreshCursor()

        If GetSuppressedEventCount() > 1 Then
            SyncLock gRefreshGridLock
                gRefreshGridScheduled = False
            End SyncLock
            EndRefreshCursor()
            Return
        End If

        If captureSelection Then
            Try
                Me.Dispatcher.Invoke(Sub() StorePendingSelection(selectionReason))
            Catch
            End Try
        End If

        Task.Run(Sub()
                     Try
                         Do
                             If GetSuppressedEventCount() = 0 Then
                                 Exit Do
                             End If

                             Thread.Sleep(200)
                         Loop

                         Me.Dispatcher.BeginInvoke(New Action(Sub()
                                                                  RefreshGrid(False, True, False)
                                                              End Sub))

                     Finally
                         ' IMPORTANT: don't clear gRefreshGridScheduled here; do it when refresh is truly finished.
                         ' We leave it to FinalizeLoad() to allow gRefreshQueued logic to work correctly.
                     End Try
                 End Sub)

    End Sub

    Public Sub ClearMonitoringOfOutlookEvents()
        Try

            If gOutlookEventHandler IsNot Nothing Then
                gOutlookEventHandler.Dispose()
                gOutlookEventHandler = Nothing
            End If

            SyncLock gListViewEntryIdsLock
                gListViewEntryIdsByFolder.Clear()
            End SyncLock

        Catch ex As Exception
        End Try

    End Sub

    Private EnsureUninteruptedProcessingOfOnEmailAddedFromEvent As New Object

    Friend Sub OnEmailAddedFromEvent(ByVal folderIndex As Integer, ByVal entryId As String, ByVal subject As String, ByVal toAddr As String, ByVal fromAddr As String, ByVal receivedTime As Date, ByVal isUnread As Boolean, ByVal body As String)

        SyncLock EnsureUninteruptedProcessingOfOnEmailAddedFromEvent

            Try

                If String.IsNullOrEmpty(entryId) OrElse folderIndex < 0 OrElse folderIndex >= gFolderTable.Length Then Return

                SetMousePointer(Cursors.Wait)

                Dim emailDetail = New StructureOfEmailDetails() With {
                    .sOriginalFolderReferenceNumber = CShort(folderIndex),
                    .sOutlookEntryID = entryId
                }

                Try

                    Dim folderInfo As FolderInfo = gFolderTable(folderIndex)

                    With emailDetail
                        .sSubject = subject
                        .sTo = toAddr
                        .sFrom = fromAddr
                        .sDateAndTime = receivedTime
                        .sUnRead = If(isUnread, System.Windows.FontWeights.Bold, System.Windows.FontWeights.Normal)
                        .sMailBoxName = GetMailboxNameFromFolderPath(folderInfo.FolderPath, folderInfo.StoreID)
                        .sBody = body
                    End With

                Catch ex As Exception

                End Try

                gEmailTableIndex += 1
                ReDim Preserve gEmailTable(gEmailTableIndex)
                gEmailTable(gEmailTableIndex) = emailDetail

                ScheduleRefreshGrid()

            Catch

            Finally
                SetMousePointer(Cursors.Arrow)
            End Try

        End SyncLock

    End Sub

#Region "Event Loop Prevention"

    Friend ReadOnly gSuppressEventLock As New Object()
    Friend ReadOnly gSuppressEventForEntryIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private _suspensionTimers As New Dictionary(Of String, System.Windows.Threading.DispatcherTimer)(StringComparer.OrdinalIgnoreCase)

    Friend Function BlockDuplicateEventProcessing(ByVal action As String, ByVal entryId As String) As Boolean

        ' The first time the routine is called for a given action + entryId it will return False
        ' Subsequent calls for the same action + entryIDToBeRemoved will return True (thus allowing the caller to suspend processing for that combination of action + entryIDToBeRemoved) 
        ' However, 1 second after having received no further calls for that action + entryIDToBeRemoved the routine will reset itself and 
        ' no longer prevent that action entryId (thus allowing processing to continue for that action + entryIDToBeRemoved based on a separate event)

        Const shortTermSuspenstionPeriodMilliseconds As Integer = 1000

        If Not Dispatcher.CheckAccess() Then
            Return Dispatcher.Invoke(Function() BlockDuplicateEventProcessing(action, entryId))
        End If

        SyncLock gSuppressEventLock

            Dim timer As System.Windows.Threading.DispatcherTimer = Nothing
            Dim alreadySuppressed As Boolean = _suspensionTimers.TryGetValue(entryId, timer)

            If alreadySuppressed Then

                timer.Stop()
                timer.Interval = TimeSpan.FromMilliseconds(shortTermSuspenstionPeriodMilliseconds)

            Else

                gSuppressEventForEntryIds.Add(entryId)

                timer = New DispatcherTimer() With {
                    .IsEnabled = False,
                    .Interval = TimeSpan.FromMilliseconds(shortTermSuspenstionPeriodMilliseconds)
                 }

                AddHandler timer.Tick, Sub(sender, e)
                                           timer.Stop()
                                           SyncLock gSuppressEventLock
                                               gSuppressEventForEntryIds.Remove(entryId)
                                               _suspensionTimers.Remove(entryId)
                                               ' Beep() ' for debugging an optional beep can be placed here to give an indication of when the suppression period ends
                                           End SyncLock
                                       End Sub

                _suspensionTimers(entryId) = timer

            End If

            timer.Start()

            Return alreadySuppressed
        End SyncLock

    End Function

#End Region

    Private EnsureUninteruptedProcessingOfOnEmailChangedFromEvent As New Object
    Friend Sub OnEmailChangedFromEvent(ByVal folderIndex As Integer, ByVal entryId As String, ByVal isUnread As Boolean, Optional ByVal attempt As Integer = 0, Optional ByVal MailItem As Microsoft.Office.Interop.Outlook.MailItem = Nothing)

        SyncLock EnsureUninteruptedProcessingOfOnEmailChangedFromEvent
            Try

                SetMousePointer(Cursors.Wait)

                Dim indexToUpdate As Integer = -1

                For i As Integer = 0 To gEmailTableIndex - 1
                    If String.IsNullOrEmpty(gEmailTable(i).sOutlookEntryID) Then
                        Continue For
                    End If
                    If gEmailTable(i).sOutlookEntryID = entryId Then
                        indexToUpdate = i
                        Exit For
                    End If
                Next

                If indexToUpdate >= 0 Then

                    Try

                        With gEmailTable(indexToUpdate)
                            If indexToUpdate < gEmailTableIndex AndAlso .sOutlookEntryID = entryId Then
                                Dim currentUnread As Boolean = (.sUnRead = System.Windows.FontWeights.Bold)
                                If currentUnread <> isUnread Then
                                    ' in a change event the only thing that should be changing is the read/unread status
                                    .sUnRead = If(isUnread, System.Windows.FontWeights.Bold, System.Windows.FontWeights.Normal)
                                End If
                            End If
                        End With

                    Catch

                    End Try

                End If

            Catch
            Finally
                ' UpdateListView() not required here as the read/unread status change will be reflected via Outlook's own UI update mechanisms
                If MailItem IsNot Nothing Then
                    Try
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(MailItem)
                    Catch ex As Exception
                    End Try
                End If
                SetMousePointer(Cursors.Arrow)
            End Try

        End SyncLock

    End Sub

    Private Class OutlookEventHandler
        Implements IDisposable

        Private ReadOnly _mainWindow As MainWindow
        Private _outlookApp As Microsoft.Office.Interop.Outlook.Application
        Private _outlookNamespace As Microsoft.Office.Interop.Outlook.NameSpace
        Private ReadOnly _storeItems As New List(Of Microsoft.Office.Interop.Outlook.Items)
        Private ReadOnly _storeItemEvents As New List(Of Microsoft.Office.Interop.Outlook.ItemsEvents_Event)
        Private Shared ReadOnly _storeRegistrationLock As New Object()
        Private Shared ReadOnly _registeredStoreIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _registeredStoreIdsForHandler As New List(Of String)()

        Public Sub New(ByVal mainWindow As MainWindow, ByVal outlookNamespace As Microsoft.Office.Interop.Outlook.NameSpace)

            Try
                _mainWindow = mainWindow
                _outlookNamespace = outlookNamespace

                _outlookApp = Nothing
                Try
                    If _outlookNamespace IsNot Nothing Then
                        _outlookApp = TryCast(_outlookNamespace.Application, Microsoft.Office.Interop.Outlook.Application)
                    End If
                Catch
                End Try

                If _outlookApp Is Nothing Then
                    Try
                        _outlookApp = New Microsoft.Office.Interop.Outlook.Application()
                    Catch
                    End Try
                End If

                If _outlookApp IsNot Nothing Then
                    If _outlookNamespace Is Nothing Then
                        Try
                            _outlookNamespace = _outlookApp.GetNamespace("MAPI")
                        Catch
                        End Try
                    End If

                    If _outlookNamespace IsNot Nothing Then

                        AddOutlookHandlers()

                    End If
                End If

            Catch ex As Exception
            End Try

        End Sub

        Private Sub AddOutlookHandlers()

            ' setup the Outlook handlers on the UI thread so that they don't get garbage collected (and stop working)

            _mainWindow.Dispatcher.Invoke(Sub()

                                              Try
                                                  If _outlookApp IsNot Nothing AndAlso Not _mainWindow.gOutlookQuitHooked Then
                                                      Dim quitEvents As Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event = TryCast(_outlookApp, Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event)
                                                      If quitEvents IsNot Nothing Then
                                                          AddHandler quitEvents.Quit, AddressOf _mainWindow.OnOutlookQuit
                                                          _mainWindow.gOutlookQuitHooked = True
                                                      End If
                                                  End If
                                              Catch
                                              End Try

                                              _TotalMailBoxes = _outlookNamespace.Stores.Count()

                                              For Each store As Microsoft.Office.Interop.Outlook.Store In _outlookNamespace.Stores
                                                  Try
                                                      Dim storeId As String = ""
                                                      Try
                                                          storeId = store.StoreID
                                                      Catch
                                                      End Try

                                                      Dim skipStore As Boolean = False
                                                      If Not String.IsNullOrEmpty(storeId) Then
                                                          SyncLock _storeRegistrationLock
                                                              If _registeredStoreIds.Contains(storeId) Then
                                                                  skipStore = True
                                                              End If
                                                          End SyncLock
                                                      End If

                                                      If skipStore Then
                                                          System.Runtime.InteropServices.Marshal.ReleaseComObject(store)
                                                          Continue For
                                                      End If

                                                      Dim inbox As Microsoft.Office.Interop.Outlook.Folder = Nothing
                                                      Dim sentMail As Microsoft.Office.Interop.Outlook.Folder = Nothing

                                                      Dim attached As Boolean = False

                                                      Try
                                                          inbox = TryCast(store.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox), Microsoft.Office.Interop.Outlook.Folder)
                                                      Catch

                                                      End Try

                                                      Try
                                                          sentMail = TryCast(store.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail), Microsoft.Office.Interop.Outlook.Folder)
                                                      Catch
                                                      End Try

                                                      Try
                                                          If inbox IsNot Nothing Then
                                                              Dim items As Microsoft.Office.Interop.Outlook.Items = inbox.Items
                                                              Dim ev As Microsoft.Office.Interop.Outlook.ItemsEvents_Event = TryCast(items, Microsoft.Office.Interop.Outlook.ItemsEvents_Event)
                                                              If ev IsNot Nothing Then
                                                                  AddHandler ev.ItemAdd, AddressOf OnItemAdd
                                                                  AddHandler ev.ItemRemove, AddressOf OnItemRemove
                                                                  AddHandler ev.ItemChange, AddressOf OnItemChange
                                                                  _storeItems.Add(items)
                                                                  _storeItemEvents.Add(ev)
                                                                  attached = True
                                                              Else
                                                                  System.Runtime.InteropServices.Marshal.ReleaseComObject(items)
                                                              End If
                                                          End If
                                                      Finally
                                                          If inbox IsNot Nothing Then
                                                              System.Runtime.InteropServices.Marshal.ReleaseComObject(inbox)
                                                              inbox = Nothing
                                                          End If
                                                      End Try

                                                      Try
                                                          If sentMail IsNot Nothing Then
                                                              Dim items As Microsoft.Office.Interop.Outlook.Items = sentMail.Items
                                                              Dim ev As Microsoft.Office.Interop.Outlook.ItemsEvents_Event = TryCast(items, Microsoft.Office.Interop.Outlook.ItemsEvents_Event)
                                                              If ev IsNot Nothing Then
                                                                  AddHandler ev.ItemAdd, AddressOf OnItemAdd
                                                                  AddHandler ev.ItemRemove, AddressOf OnItemRemove
                                                                  AddHandler ev.ItemChange, AddressOf OnItemChange
                                                                  _storeItems.Add(items)
                                                                  _storeItemEvents.Add(ev)
                                                                  attached = True
                                                              Else
                                                                  System.Runtime.InteropServices.Marshal.ReleaseComObject(items)
                                                              End If
                                                          End If
                                                      Finally
                                                          If sentMail IsNot Nothing Then
                                                              System.Runtime.InteropServices.Marshal.ReleaseComObject(sentMail)
                                                              sentMail = Nothing
                                                          End If
                                                      End Try

                                                      If attached AndAlso Not String.IsNullOrEmpty(storeId) Then
                                                          SyncLock _storeRegistrationLock
                                                              If Not _registeredStoreIds.Contains(storeId) Then
                                                                  _registeredStoreIds.Add(storeId)
                                                              End If
                                                              _registeredStoreIdsForHandler.Add(storeId)
                                                          End SyncLock
                                                      End If

                                                  Catch ex As Exception
                                                  Finally
                                                      Try
                                                          System.Runtime.InteropServices.Marshal.ReleaseComObject(store)
                                                      Catch
                                                      End Try
                                                  End Try
                                              Next
#If DEBUG Then
                                              Console.WriteLine("Outlook event handlers attached.")
#End If
                                          End Sub)

        End Sub

        Private ReadOnly EnsureUninteruptedProcessingOfOnItemAdd As New Object
        Private Sub OnItemAdd(ByVal Item As Object)

            If gIsRefreshing Then Exit Sub

            SyncLock EnsureUninteruptedProcessingOfOnItemAdd

                Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = Nothing
                Dim folder As Outlook.MAPIFolder = Nothing

                Try

                    mailItem = TryCast(Item, Microsoft.Office.Interop.Outlook.MailItem)

                    If mailItem Is Nothing Then Return
                    If mailItem.EntryID Is Nothing Then Return

                    Dim entryId As String = ""

                    Try
                        entryId = mailItem.EntryID
                    Catch
                        Return
                    End Try

                    If _mainWindow.BlockDuplicateEventProcessing("Add", entryId) Then
                        Return
                    End If

                    Dim folderIdx As Integer = -1
                    Dim subject As String = ""
                    Dim toAddr As String = ""
                    Dim fromAddr As String = ""
                    Dim receivedTime As Date = Nothing
                    Dim isUnread As Boolean = False
                    Dim body As String = ""

                    Try
                        folder = CType(mailItem.Parent, Outlook.MAPIFolder)

                        folderIdx = 0
                        For Each entry In gFolderTable
                            If String.Equals(entry.FolderPath, folder.FolderPath, StringComparison.OrdinalIgnoreCase) Then
                                Exit For
                            End If
                            folderIdx += 1
                        Next

                        If folderIdx >= gFolderTable.Length Then
                            folderIdx = -1
                        End If
                    Catch
                        folderIdx = -1
                    End Try

                    Try
                        subject = _mainWindow.CleanUpSubjectLine(If(mailItem.Subject, String.Empty))
                        toAddr = If(mailItem.To, String.Empty)
                        receivedTime = mailItem.ReceivedTime
                        isUnread = mailItem.UnRead
                        body = If(mailItem.Body, String.Empty)

                        Dim friendlyFrom As String = ""
                        Try
                            If mailItem.Sender IsNot Nothing Then
                                If String.Equals(mailItem.SenderEmailType, "SMTP", StringComparison.OrdinalIgnoreCase) Then
                                    friendlyFrom = mailItem.SenderEmailAddress
                                Else
                                    Dim exUser As Microsoft.Office.Interop.Outlook.ExchangeUser =
                                        TryCast(mailItem.Sender.GetExchangeUser(), Microsoft.Office.Interop.Outlook.ExchangeUser)
                                    If exUser IsNot Nothing AndAlso Not String.IsNullOrEmpty(exUser.PrimarySmtpAddress) Then
                                        friendlyFrom = exUser.PrimarySmtpAddress
                                    Else
                                        friendlyFrom = mailItem.SenderEmailAddress
                                    End If
                                End If
                            Else
                                friendlyFrom = mailItem.SenderEmailAddress
                            End If
                        Catch
                            friendlyFrom = mailItem.SenderEmailAddress
                        End Try

                        fromAddr = If(friendlyFrom, String.Empty)
                    Catch
                    End Try

                    If folder IsNot Nothing Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(folder)
                        folder = Nothing
                    End If

                    If mailItem IsNot Nothing Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                        mailItem = Nothing
                    End If

                    If folderIdx < 0 Then Return

                    _mainWindow.Dispatcher.BeginInvoke(New Action(Sub()
                                                                      _mainWindow.OnEmailAddedFromEvent(folderIdx, entryId, subject, toAddr, fromAddr, receivedTime, isUnread, body)
                                                                  End Sub))

                Catch ex As Exception

                Finally
                    Try
                        If folder IsNot Nothing Then
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(folder)
                        End If
                    Catch
                    End Try

                    Try
                        If mailItem IsNot Nothing Then
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                        End If
                    Catch
                    End Try

                End Try

            End SyncLock

        End Sub

        Private Sub OnItemRemove()

            If gIsRefreshing Then Exit Sub

            ' OnItemRemove can be triggered in two scenarios:
            ' 1. the email is deleted or removed from a monitored folder via FileFriendly
            ' 2. the email is deletes or removed from a monitored folder via Outlook

            ' However, in neither scenario do we know the EntryID of the deleted / removed email
            ' I had originally tried to track the EntryIDs in the listview by taking a snapshot of them before and after the removal and then comparing the two snapshots to deduce which EntryID(s) had been removed
            ' However, this approach does not work in second case (removal in Outlook) as the listview doesn't change just because an email is deleted in Outlook

            ' Accordingly, on removal the program will simply refresh the grid view to pick up any changes

            _mainWindow.ScheduleRefreshGrid()

        End Sub

        Private Sub OnItemChange(ByVal Item As Object)

            ' OnItemChange being called when an email is marked read/unread in Outlook
            ' This event can be triggered by many different changes to an email item (for example changing its priority)
            ' we are only interested in processing read/unread changes here

            If gIsRefreshing Then Exit Sub

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = Nothing

            Try

                mailItem = TryCast(Item, Microsoft.Office.Interop.Outlook.MailItem)
                If mailItem Is Nothing Then Return
                If mailItem.EntryID Is Nothing Then Return

                ' Block duplicate requests for the same email
                Dim entryId As String = mailItem.EntryID
                Dim action As String = If(mailItem.UnRead, "ReadGoingToUnread", "UnreadGoingToRead")
                If _mainWindow.BlockDuplicateEventProcessing(action, entryId) Then Return

                _mainWindow.ScheduleRefreshGrid()

            Catch ex As Exception

            End Try

        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Try

                If _outlookApp IsNot Nothing Then
                    Try
                        _mainWindow.Dispatcher.Invoke(Sub()
                                                          Try
                                                              If _mainWindow.gOutlookQuitHooked Then
                                                                  Dim quitEvents As Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event = TryCast(_outlookApp, Microsoft.Office.Interop.Outlook.ApplicationEvents_11_Event)
                                                                  If quitEvents IsNot Nothing Then
                                                                      RemoveHandler quitEvents.Quit, AddressOf _mainWindow.OnOutlookQuit
                                                                  End If
                                                                  _mainWindow.gOutlookQuitHooked = False
                                                              End If
                                                          Catch
                                                          End Try
                                                      End Sub)
                    Catch
                    End Try
                End If

                If _storeItemEvents IsNot Nothing AndAlso _storeItems IsNot Nothing Then
                    For i As Integer = 0 To _storeItemEvents.Count - 1
                        Dim ev As Microsoft.Office.Interop.Outlook.ItemsEvents_Event = _storeItemEvents(i)
                        Dim items As Microsoft.Office.Interop.Outlook.Items = _storeItems(i)

                        If ev IsNot Nothing Then
                            RemoveHandler ev.ItemAdd, AddressOf OnItemAdd
                            RemoveHandler ev.ItemRemove, AddressOf OnItemRemove
                            RemoveHandler ev.ItemChange, AddressOf OnItemChange
                        End If

                        If items IsNot Nothing Then
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(items)
                        End If
                    Next
#If DEBUG Then
                    Console.WriteLine("Outlook event handlers detached.")
#End If

                End If

                SyncLock _storeRegistrationLock
                    For Each storeId In _registeredStoreIdsForHandler
                        _registeredStoreIds.Remove(storeId)
                    Next
                End SyncLock
                _registeredStoreIdsForHandler.Clear()

                _storeItemEvents.Clear()
                _storeItems.Clear()

                If _outlookNamespace IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_outlookNamespace)
                    _outlookNamespace = Nothing
                End If

                If _outlookApp IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_outlookApp)
                    _outlookApp = Nothing
                End If

            Catch ex As Exception
            End Try
        End Sub

    End Class

    Private Class IntegerComparer
        Implements IEqualityComparer(Of Integer)

        Public Shared ReadOnly Instance As New IntegerComparer()

        Public Overrides Function Equals(obj As Object) As Boolean
            Return MyBase.Equals(obj)
        End Function

        Public Overloads Function Equals(x As Integer, y As Integer) As Boolean Implements IEqualityComparer(Of Integer).Equals
            Return x = y
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return MyBase.GetHashCode()
        End Function

        Public Overloads Function GetHashCode(obj As Integer) As Integer Implements IEqualityComparer(Of Integer).GetHashCode
            Return obj.GetHashCode()
        End Function
    End Class

#End Region

End Class

Public Class EMailTableSorter

    Implements System.Collections.IComparer

    Public Enum SortOrder As Integer
        None = 0
        Ascending = 1
        Descending = 2
    End Enum

    Private ReadOnly ObjectCompare As Comparer 'CaseInsensitiveComparer

    Private PrimaryColumnToSort As Integer
    Private PrimaryOrderOfSort As SortOrder
    Private SecondaryColumnToSort As Integer
    Private SecondaryOrderOfSort As SortOrder

    Public Property PrimarySortColumn() As Integer
        Set(ByVal Value As Integer)
            PrimaryColumnToSort = Value
        End Set
        Get
            Return PrimaryColumnToSort
        End Get
    End Property

    Public Property PrimaryOrder() As SortOrder
        Set(ByVal Value As SortOrder)
            PrimaryOrderOfSort = Value
        End Set
        Get
            Return PrimaryOrderOfSort
        End Get
    End Property

    Public Property SecondarySortColumn() As Integer
        Set(ByVal Value As Integer)
            SecondaryColumnToSort = Value
        End Set
        Get
            Return SecondaryColumnToSort
        End Get
    End Property

    Public Property SecondaryOrder() As SortOrder
        Set(ByVal Value As SortOrder)
            SecondaryOrderOfSort = Value
        End Set
        Get
            Return SecondaryOrderOfSort
        End Get
    End Property

    Public Sub New()

        PrimaryColumnToSort = -1
        PrimaryOrderOfSort = SortOrder.None
        SecondaryColumnToSort = -1
        SecondaryOrderOfSort = SortOrder.None
        ObjectCompare = New Comparer(System.Globalization.CultureInfo.CurrentCulture)

    End Sub

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare

        Dim compareResult As Integer


        compareResult = ObjectCompare.Compare(x.sSubject, y.sSubject)

        If compareResult <> 0 Then

            If (PrimaryOrderOfSort = SortOrder.Ascending) Then
                Return compareResult

            ElseIf (PrimaryOrderOfSort = SortOrder.Descending) Then
                Return (-compareResult)

            Else
                Return 0

            End If

        Else

            compareResult = ObjectCompare.Compare(x.sTrailer, y.sTrailer)
            If compareResult <> 0 Then

                If (PrimaryOrderOfSort = SortOrder.Ascending) Then
                    Return compareResult

                ElseIf (PrimaryOrderOfSort = SortOrder.Descending) Then
                    Return (-compareResult)

                Else
                    Return 0

                End If

            Else

                If SecondaryColumnToSort >= 0 Then

                    compareResult = ObjectCompare.Compare(x.sDateAndTime, y.sDateAndTime)

                    If (SecondaryOrderOfSort = SortOrder.Ascending) Then
                        Return compareResult

                    ElseIf (SecondaryOrderOfSort = SortOrder.Descending) Then
                        Return (-compareResult)

                    Else
                        Return 0

                    End If

                End If

            End If

        End If


    End Function

End Class

#Region "MyFormatter"

Public Class MyFormatter
    Implements IValueConverter

    Public Function Convert(ByVal value As Object,
                            ByVal targetType As System.Type,
                            ByVal parameter As Object,
                            ByVal culture As System.Globalization.CultureInfo) As Object _
                            Implements System.Windows.Data.IValueConverter.Convert

        ' Ensure we are working with a Date value if possible
        Dim dt As Nullable(Of Date) = Nothing
        If TypeOf value Is Date Then
            dt = CType(value, Date)
        ElseIf TypeOf value Is Nullable(Of Date) Then
            dt = CType(value, Nullable(Of Date))
        ElseIf value IsNot Nothing AndAlso IsDate(value) Then
            dt = CDate(value)
        End If

        If parameter IsNot Nothing AndAlso dt.HasValue Then
            Select Case CStr(parameter)
                Case "Date"
                    ' Date column: date only
                    Return Format(dt.Value, gPreferredDateFormat)
                Case "Time"
                    ' Time column: time only
                    Return Format(dt.Value, gPreferredTimeFormat)
            End Select
        End If

        ' Fallback – no special formatting
        Return value
    End Function

    Public Function ConvertBack(ByVal value As Object,
                                ByVal targetType As System.Type,
                                ByVal parameter As Object,
                                ByVal culture As System.Globalization.CultureInfo) As Object _
                                Implements System.Windows.Data.IValueConverter.ConvertBack

        If targetType Is GetType(Date) OrElse targetType Is GetType(Nullable(Of Date)) Then
            If IsDate(value) Then
                Return CDate(value)
            ElseIf value.ToString() = "" Then
                Return Nothing
            Else
                Return Now() 'invalid type was entered so just give a default.
            End If
        ElseIf targetType Is GetType(Decimal) Then
            If IsNumeric(value) Then
                Return CDec(value)
            Else
                Return 0
            End If
        End If

        Return value
    End Function

End Class

#End Region

Public Class FinalRecommendationTableSorter

    Implements System.Collections.IComparer

    Public Enum MySortOrder As Integer
        Ascending = 1
        Descending = 2
    End Enum

    Private ReadOnly ObjectCompare As Comparer 'CaseInsensitiveComparer

    Private _PrimaryColumnToSort As String
    Private _SortOrder As MySortOrder

    Public Property PrimaryColumnToSort() As String
        Set(ByVal Value As String)
            _PrimaryColumnToSort = Value
        End Set
        Get
            Return _PrimaryColumnToSort
        End Get
    End Property

    Public Property SortOrder() As MySortOrder
        Set(ByVal Value As MySortOrder)
            _SortOrder = Value
        End Set
        Get
            Return _SortOrder
        End Get
    End Property

    Public Sub New()

        ObjectCompare = New Comparer(System.Globalization.CultureInfo.CurrentCulture)

    End Sub

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare

        Dim compareResult As Integer

        Select Case PrimaryColumnToSort

            Case Is = "MailBoxName", "Mailbox"

                ' Primary: MailBoxName; Secondary: Subject+Trailer; Tertiary: DateTime (desc)
                If SortOrder = MySortOrder.Ascending Then
                    compareResult = ObjectCompare.Compare(x.MailBoxName, y.MailBoxName)
                Else
                    compareResult = -ObjectCompare.Compare(x.MailBoxName, y.MailBoxName)
                End If

                If compareResult = 0 Then
                    ' Same mailbox: compare by Subject+Trailer ascending
                    compareResult = ObjectCompare.Compare(x.Subject & x.Trailer, y.Subject & y.Trailer)
                End If

                If compareResult = 0 Then
                    ' Same subject: compare by DateTime, most recent first
                    compareResult = ObjectCompare.Compare(x.DateTime, y.DateTime)
                    If compareResult <> 0 Then compareResult = -compareResult
                End If

            Case Is = "Subject"

                ' Primary: Subject+Trailer; Secondary: DateTime (desc)
                If SortOrder = MySortOrder.Ascending Then
                    compareResult = ObjectCompare.Compare(x.Subject & x.Trailer, y.Subject & y.Trailer)
                Else
                    compareResult = -ObjectCompare.Compare(x.Subject & x.Trailer, y.Subject & y.Trailer)
                End If

                If compareResult = 0 Then
                    compareResult = ObjectCompare.Compare(x.DateTime, y.DateTime)
                    If compareResult <> 0 Then compareResult = -compareResult
                End If

            Case Is = "To"

                If SortOrder = MySortOrder.Ascending Then
                    'sort by subject (in ascending order) and then if the subjects are the same by date (in descending order)
                    compareResult = ObjectCompare.Compare(x.xTo, y.xTo)
                    If compareResult = 0 Then
                        compareResult = ObjectCompare.Compare(x.DateTime, y.DateTime)
                        If compareResult <> 0 Then compareResult = -compareResult ' (use - for descending sort order)
                    End If

                Else

                    compareResult = -ObjectCompare.Compare(x.xTo, y.xTo) ' (use - for descending sort order
                    If compareResult = 0 Then
                        compareResult = ObjectCompare.Compare(x.DateTime, y.DateTime)
                        If compareResult <> 0 Then compareResult = -compareResult ' (use - for descending sort order)
                    End If

                End If

            Case Is = "From"

                If SortOrder = MySortOrder.Ascending Then
                    'sort by subject (in ascending order) and then if the subjects are the same by date (in descending order)
                    compareResult = ObjectCompare.Compare(x.From, y.From)
                    If compareResult = 0 Then
                        compareResult = ObjectCompare.Compare(x.DateTime, y.DateTime)
                        If compareResult <> 0 Then compareResult = -compareResult ' (use - for descending sort order)
                    End If

                Else

                    compareResult = -ObjectCompare.Compare(x.From, y.From) ' (use - for descending sort order
                    If compareResult = 0 Then
                        compareResult = ObjectCompare.Compare(x.DateTime, y.DateTime)
                        If compareResult <> 0 Then compareResult = -compareResult ' (use - for descending sort order)
                    End If

                End If


            Case "Date", "Time"

                If SortOrder = MySortOrder.Ascending Then
                    'sort by subject (in ascending order) and then if the subjects are the same by date (in descending order)
                    compareResult = ObjectCompare.Compare(x.DateTime, y.DateTime)
                    If compareResult = 0 Then
                        compareResult = ObjectCompare.Compare(x.Subject & x.Trailer, y.Subject & y.Trailer)
                        If compareResult <> 0 Then compareResult = -compareResult ' (use - for descending sort order)
                    End If

                Else

                    compareResult = -ObjectCompare.Compare(x.DateTime, y.DateTime) ' (use - for descending sort order
                    If compareResult = 0 Then
                        compareResult = ObjectCompare.Compare(x.Subject & x.Trailer, y.Subject & y.Trailer)
                        If compareResult <> 0 Then compareResult = -compareResult ' (use - for descending sort order)
                    End If

                End If

        End Select

        Return compareResult

    End Function

End Class