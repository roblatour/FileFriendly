Imports System.Linq
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
'Imports System.Windows.Forms
Imports System.Windows.Threading
Imports Microsoft.Office.Interop
'Imports Microsoft.Office.Interop.Outlook

Class MainWindow
    Inherits Window

    Private gForegroundColourAlert As System.Windows.Media.SolidColorBrush
    Private gForegroundColourEnabled As System.Windows.Media.SolidColorBrush
    Private gForegroundColourDisabled As System.Windows.Media.SolidColorBrush

    Private gProgressUpdateTimer As System.Threading.Timer
    Private ReadOnly gProgressUpdateTimerLock As New Object
    Private gProgressCounter As Long
    Private gSuppressUpdatesToDetailBox As Boolean = False

    Enum ActionType As Integer
        None = 1
        Hide = 1
        Delete = 2
        File = 3
        ToggleRead = 4
    End Enum
    Structure StructureOfUndoLog
        Dim FixedIndex As Integer
        Dim ActionApplied As ActionType
        Dim EmailEntryID As String
        Dim SourceStoreID As String
        Dim SourceFolderEntryID As String
        Dim TargetStoreID As String
        Dim TargetFolderEntryID As String
        Dim LvrcItem As ListViewRowClass
    End Structure

    Private Enum SortOrder As Integer
        None = 0
        Ascending = 1
        Descending = 2
    End Enum
    Structure StructureOfEmailDetails
        Dim sOutlookEntryID As String
        Dim sSubject As String
        Dim sDateAndTime As DateTime
        Dim sTo As String
        Dim sFrom As String
        Dim sOriginalFolderReferenceNumber As Integer
        Dim sRecommendedFolder1ReferenceNumber As Integer
        Dim sRecommendedFolder2ReferenceNumber As Integer
        Dim sRecommendedFolder3ReferenceNumber As Integer
        Dim sRecommendedFolderFinalReferenceNumber As Integer
        Dim sUnRead As FontWeight
        Dim sMailBoxName As String ' mailbox/postbox name
        Dim sTrailer As String
    End Structure

    Friend gEmailTable(1) As StructureOfEmailDetails
    Private gEmailTableIndex As Integer = 0
    Private Const gEmailTableGrowth As Integer = 200 ' when more space is needed, grow the table by this many entries

    Structure StructureOfLastComputedTrailerTable
        Dim sOutlookEntryID As String
        Dim sTrailer As String
    End Structure

    Private lTotalEMails As Integer = 0
    Private lTotalEMailsToBeReviewed As Integer = 0
    Private lTotalRecommendations As Integer = 0

    Private UniqueSubjectsMap As New Dictionary(Of String, Dictionary(Of Integer, Integer))(StringComparer.Ordinal)

    Private gOriginalWidthSubject, gOriginalWidthTo, gOriginalWidthFrom, gOriginalWidthDate As Integer

    Private gViewSent As Boolean = True
    Private gViewInbox As Boolean = True
    Private gViewOtherFolders As Boolean = True
    Private gViewRead As Boolean = True
    Private gViewUnRead As Boolean = True

    Private gFinalRecommendationTable(1) As ListViewRowClass

    Friend gHiddenEntryIDs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly gHiddenEntryIdsFilePath As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FileFriendly", "ListOfEmailIdentifiersToKeepHidden.txt")

    Private Shared gIsRefreshing As Boolean = False
    Private ReadOnly gRefreshGateLock As New Object

    Private Shared gCancelRefresh As Boolean = False

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
    'Private gInboxFolderIndices As New List(Of Integer)
    'Private gSentFolderIndices As New List(Of Integer)

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
        Hide = 0
        OutlookDelete = 1
        Refresh = 2
        Sort = 3
        ToggleRead = 4
        Undo = 5
        UserDelete = 6
    End Enum

    Private Enum ListSortDirection
        Ascending = 1
        Descending = 2
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

    Private gClosingNow As Boolean = False
    Private CloseWindow As New System.Windows.Forms.MethodInvoker(AddressOf CloseWindowNow)
    Private MoveMainWindow As New System.Windows.Forms.MethodInvoker(AddressOf MoveMainWindowNow)
    Private UpdateContextMenu As New System.Windows.Forms.MethodInvoker(AddressOf UpdateContextMenuNow)
    Private ResizeMainWindow As New System.Windows.Forms.MethodInvoker(AddressOf ResizeMainWindowNow)
    Private UpdateRecommendation As New System.Windows.Forms.MethodInvoker(AddressOf UpdateRecommendationNow)
    Private ReadOnly gRefreshCursorLock As New Object()
    Private gRefreshCursorRefCount As Integer = 0
    Private gRefreshQueued As Boolean = False
    Private Collection_of_folders_to_exclude = New System.Collections.Specialized.StringCollection
    Private Collection_of_folders_to_exclude_is_empty As Boolean = True
    Private lBlankEMailDetailRecord As StructureOfEmailDetails
    Private lWhenSent As Boolean
    Private ActivateMenu As New System.Windows.Forms.MethodInvoker(AddressOf ActivateMenuNow)
    Private PerformActionByProxy As New System.Windows.Forms.MethodInvoker(AddressOf PerformActionByProxyNow)
    Private UpdateReadToggleContextMenu As New System.Windows.Forms.MethodInvoker(AddressOf UpdateReadToggleContextMenuNow)
    Private SelectedListViewItem As New ListViewRowClass
    Private gUndoLogIndex As Integer = 0
    Private gUndoLogSubIndex As Integer = 0
    Private gUndoLogMaxEntries As Integer = 500
    Private gUndoLogMaxSubEntries As Integer = 500
    Private gUndoLog(gUndoLogMaxEntries, gUndoLogMaxSubEntries) As StructureOfUndoLog
    Private gUndoLogWasUpdated As Boolean = False
    Private _lastDirection As ListSortDirection = ListSortDirection.Descending
    Private _lastheaderClicked As GridViewColumnHeader
    Friend gCurrentSortOrder As String = "Mailbox"
    Private gCurrentSortDirection As ListSortDirection = ListSortDirection.Ascending
    Private gOutlookEventHandler As OutlookEventHandler
    Private ReadOnly gEventHandlerLock As New Object()
    Private gRefreshGridScheduled As Boolean = False
    Private ReadOnly gRefreshGridLock As New Object()
    Private EnsureUninteruptedProcessingOfOnEmailAddedFromEvent As New Object
    Friend ReadOnly gSuppressEventLock As New Object()
    Friend ReadOnly gSuppressEventForEntryIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private _suspensionTimers As New Dictionary(Of String, System.Windows.Threading.DispatcherTimer)(StringComparer.OrdinalIgnoreCase)
    Private Shared ReadOnly SubjectPrefixRegex As New Regex("^(?:\s*(?:RE|FW):\s*)+", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public intRecommendationFinal As String = ""

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.        MainWindow.Visibility = Windows.Visibility.Visible
        EnsureOnlyOneInstanceOfApp()

        gMainWindow = Me

        SetProcessPriorities("Initialize")

    End Sub

    Public Sub SafelyMoveMainWindow()
        Call Dispatcher.BeginInvoke(MoveMainWindow)
    End Sub

    Public Sub SafelyUpdateContextMenu()
        Call Dispatcher.BeginInvoke(UpdateContextMenu)
    End Sub

    Public Sub SafelyResizeMainWindow()
        Call Dispatcher.BeginInvoke(ResizeMainWindow)
    End Sub

    Public Sub SafelyUpdateRecommendationFromPickAFolderWindow()
        Call Dispatcher.BeginInvoke(UpdateRecommendation)
    End Sub

    Public Function CleanUpSubjectLine(subjectLine As String) As String

        ' remove all "RE:", "FWD:", and "FW:" prefixes from the subject line

        Static SubjectPrefixRegex As New Regex("^(?:\s*(?:RE|FW(?:D)?):\s*)+", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

        Dim result = SubjectPrefixRegex.Replace(subjectLine, "").Trim()
        Return If(String.IsNullOrWhiteSpace(result), " ", result)

    End Function

    Public Sub SafelyActivateMenu()
        Call Dispatcher.BeginInvoke(ActivateMenu)
    End Sub

    Public Sub SafelyPerformActionByProxy()
        Call Dispatcher.BeginInvoke(PerformActionByProxy)
    End Sub

    Public Sub SafelyUpdateReadToggleContextMenu()
        Call Dispatcher.BeginInvoke(UpdateReadToggleContextMenu)
    End Sub

    Public Function FileMessage(ByVal oldEmailEntryID As String,
                                ByVal SourceStoreID As String,
                                ByVal SourceFolderEntryID As String,
                                ByVal TargetStoreID As String,
                                ByVal TargetFolderEntryID As String) As String

        Dim ReturnCode As String = ""

        Dim mail As Microsoft.Office.Interop.Outlook.MailItem = Nothing
        Dim targetFolder As Microsoft.Office.Interop.Outlook.MAPIFolder = Nothing
        Dim oMovedEmail As Microsoft.Office.Interop.Outlook.MailItem = Nothing

        Try

            ' Retry GetItemFromID / GetFolderFromID once on RPC disconnect
            Dim attempt As Integer = 0
            While attempt < 2 AndAlso (mail Is Nothing OrElse targetFolder Is Nothing)

                Try
                    mail = TryCast(oNS.GetItemFromID(oldEmailEntryID, SourceStoreID), Microsoft.Office.Interop.Outlook.MailItem)
                    targetFolder = oNS.GetFolderFromID(TargetFolderEntryID, TargetStoreID)

                Catch comEx As System.Runtime.InteropServices.COMException
                    Const RPC_E_SERVER_UNAVAILABLE As Integer = &H800706BA
                    Const RPC_E_DISCONNECTED As Integer = &H800706BE

                    If (comEx.HResult = RPC_E_SERVER_UNAVAILABLE OrElse comEx.HResult = RPC_E_DISCONNECTED) AndAlso attempt = 0 Then
                        oNS = Nothing
                        oApp = Nothing
                        mail = Nothing
                        targetFolder = Nothing
                    Else

                        Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

                        Call ShowMessageBox("FileFriendly - Outlook Error",
                             CustomDialog.CustomDialogIcons.Information,
                             "Unexpected Error!",
                             "FileFriendly has encountered an unexpected Error." & vbCrLf & "FileFriendly could not complete the requested action In Outlook. (1)",
                             currentMethodName & " - " & comEx.ToString,
                             "",
                             CustomDialog.CustomDialogIcons.None,
                             CustomDialog.CustomDialogButtons.OK,
                             CustomDialog.CustomDialogResults.OK)

                        Return ""
                    End If

                Catch ex As Exception

                    Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

                    Call ShowMessageBox("FileFriendly - Outlook Error",
                         CustomDialog.CustomDialogIcons.Information,
                         "Unexpected Error!",
                         "FileFriendly has encountered an unexpected Error." & vbCrLf & "FileFriendly could not complete the requested action In Outlook. (2)",
                         currentMethodName & " - " & ex.ToString,
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

            ' the move being done below will itself raise a 'Remove' event which we need to ignore ' ---------- the remove now cause a grid refresh so we don't need to block it here
            _MainWindow.BlockDuplicateEventProcessing("Remove", oldEmailEntryID) ' used to start the suppression of the removes event that will be raised by Outlook due to the move

            'Do the move

            oMovedEmail = mail.Move(targetFolder)

            'Get new Entry ID
            ReturnCode = oMovedEmail.EntryID

        Catch ex As Exception

            Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

            Call ShowMessageBox("FileFriendly - Outlook Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected Error.",
                 currentMethodName & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

            ReturnCode = ""

        Finally

            If oMovedEmail IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oMovedEmail)
                oMovedEmail = Nothing
            End If

            If mail IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(mail)
                mail = Nothing
            End If

            If targetFolder IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(targetFolder)
                targetFolder = Nothing
            End If

        End Try

        Return ReturnCode

    End Function

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

    Friend Sub OnEmailAddedFromEvent(ByVal folderIndex As Integer, ByVal entryId As String, ByVal subject As String, ByVal toAddr As String, ByVal fromAddr As String, ByVal receivedTime As Date, ByVal isUnread As Boolean, ByVal body As String)

        SyncLock EnsureUninteruptedProcessingOfOnEmailAddedFromEvent

            Try

                If String.IsNullOrEmpty(entryId) OrElse folderIndex < 0 OrElse folderIndex >= gFolderTable.Length Then Return

                SetMousePointer(Cursors.Wait)

                ' add the new email to the email table

                Dim emailDetail = New StructureOfEmailDetails() With {
                    .sOriginalFolderReferenceNumber = folderIndex,
                    .sOutlookEntryID = entryId
                }

                gEmailTableIndex += 1
                ReDim Preserve gEmailTable(gEmailTableIndex)

                Dim folderInfo As FolderInfo = gFolderTable(folderIndex)

                With emailDetail
                    .sSubject = subject
                    .sTo = toAddr
                    .sFrom = fromAddr
                    .sDateAndTime = receivedTime
                    .sUnRead = If(isUnread, System.Windows.FontWeights.Bold, System.Windows.FontWeights.Normal)
                    .sMailBoxName = GetMailboxNameFromFolderPath(folderInfo.FolderPath, folderInfo.StoreID)
                    .sTrailer = CreateTrailer(.sDateAndTime, subject, body)
                End With


                gEmailTable(gEmailTableIndex) = emailDetail

                ScheduleRefreshGrid()

            Catch ex As Exception

#If DEBUG Then
                Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name
                Console.WriteLine(currentMethodName & " - " & ex.ToString)
#End If

            Finally
                SetMousePointer(Cursors.Arrow)
            End Try

        End SyncLock

    End Sub

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

    Private Sub LiftSuspensionOnRemovedEntryId(ByVal entryId As String)

        SyncLock gSuppressEventLock

            Try
                gSuppressEventForEntryIds.Remove(entryId)
                _suspensionTimers.Remove(entryId)
            Catch
            End Try

        End SyncLock

    End Sub

    Private Sub SafelyCloseWindow()
        Call Dispatcher.BeginInvoke(CloseWindow)
    End Sub
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

            Randomize()

            LoadHiddenEntryIds()

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

                GracefulShutdown()
                Application.Current.Shutdown()
                Return
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

            'set screen width and height (managing the case where the client has changed screen resolutions _from last run)

            '****** width
            Dim dCurrentScreenWidth As Double = System.Windows.SystemParameters.PrimaryScreenWidth
            If My.Settings.ScreenWidth = dCurrentScreenWidth Then
                ' no need to change settings _from last time
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

            Dim mainPlacement As System.Windows.Rect = AdjustWindowRect(My.Settings.MainLeft, My.Settings.MainTop, My.Settings.MainWidth, My.Settings.MainHeight, Me.MinWidth, Me.MinHeight)

            Me.Width = mainPlacement.Width
            Me.Height = mainPlacement.Height
            Me.Left = mainPlacement.X
            Me.Top = mainPlacement.Y

            gmwWidth = Me.ActualWidth
            gmwHeight = Me.ActualHeight
            gmwLeft = Me.Left
            gmwTop = Me.Top
            PAFWSaysMWLeftShouldBe = gmwLeft
            PAFWSaysMWTopShouldBe = gmwTop

            Dim dCurrentScreenHeight As Double = System.Windows.SystemParameters.PrimaryScreenHeight
            If My.Settings.ScreenHeight = dCurrentScreenHeight Then
                ' no need to change settings _from last time
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

    Private Sub LoadHiddenEntryIds()

        gHiddenEntryIDs.Clear()

        If Not My.Settings.KeepHiddenEmailsHidden Then
            Return
        End If

        Try
            If System.IO.File.Exists(gHiddenEntryIdsFilePath) Then
                Dim lines = System.IO.File.ReadAllLines(gHiddenEntryIdsFilePath)
                For Each line In lines
                    Dim trimmed = line.Trim()
                    If trimmed.Length > 0 Then
                        gHiddenEntryIDs.Add(trimmed)
                    End If
                Next
            End If
        Catch
            gHiddenEntryIDs.Clear()
        End Try

    End Sub

    Private Sub SaveHiddenEntryIds()

        If Not My.Settings.KeepHiddenEmailsHidden Then
            Try
                If System.IO.File.Exists(gHiddenEntryIdsFilePath) Then
                    System.IO.File.Delete(gHiddenEntryIdsFilePath)
                End If
            Catch
            End Try
            Return
        End If

        Try
            Dim directory = System.IO.Path.GetDirectoryName(gHiddenEntryIdsFilePath)
            If Not System.IO.Directory.Exists(directory) Then
                System.IO.Directory.CreateDirectory(directory)
            End If
            System.IO.File.WriteAllLines(gHiddenEntryIdsFilePath, gHiddenEntryIDs)
        Catch
        End Try

    End Sub

    Private Sub GracefulShutdown()

        On Error Resume Next

        SetProcessPriorities("Shutdown")

        SaveHiddenEntryIds()

        My.Settings.MainWidth = Me.ActualWidth
        My.Settings.MainHeight = Me.ActualHeight
        My.Settings.MainTop = Me.Top
        My.Settings.MainLeft = Me.Left

        If gPickAFolderWindow IsNot Nothing Then
            My.Settings.FoldersWidth = gPickAFolderWindow.ActualWidth
            My.Settings.FoldersHeight = gPickAFolderWindow.ActualHeight
            My.Settings.FoldersTop = gPickAFolderWindow.Top
            My.Settings.FoldersLeft = gPickAFolderWindow.Left
        End If

        My.Settings.StartDocked = gWindowDocked

        'this should always be true, but check anyway
        If System.Windows.SystemParameters.PrimaryScreenWidth > 0 Then
            My.Settings.ScreenWidth = System.Windows.SystemParameters.PrimaryScreenWidth
        End If

        If System.Windows.SystemParameters.PrimaryScreenHeight > 0 Then
            My.Settings.ScreenHeight = System.Windows.SystemParameters.PrimaryScreenHeight
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

    Private Sub UpdateRecommendationNow()

        PerformAction("File", False)
        gPickFromContextMenuOverride = -1

    End Sub

    Private Sub SetProcessPriorities(ByVal Command As String)

        Static Dim myProcess, OutlookProcess As Process

        Try

            Select Case Command

                Case Is = "Initialize"

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

    ' Thread‑safe wrapper to update the cursor _from any thread
    Private Sub SetMousePointer(ByVal cursor As System.Windows.Input.Cursor)
        Dim updateCursor As Action = Sub()
                                         Me.Cursor = cursor
                                         If Me.ListView1 IsNot Nothing Then
                                             Me.ListView1.Cursor = cursor
                                         End If
                                         If gPickAFolderWindow IsNot Nothing Then
                                             gPickAFolderWindow.Cursor = cursor
                                         End If
                                     End Sub

        If Dispatcher.CheckAccess() Then
            updateCursor()
        Else
            Dispatcher.BeginInvoke(updateCursor)
        End If
    End Sub

    ' Keeps the Wait cursor active across scheduled + running refresh work.

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
        MenuOptionEnabled("ToggleRead", False)

        gIsRefreshing = True
        gCancelRefresh = False

        UpdateRefreshMenuState()

        gViewInbox = gRefreshInbox
        gViewSent = gRefreshSent
        gViewOtherFolders = gRefreshOtherFolders

        MenuViewInbox.IsChecked = gRefreshInbox
        MenuViewSent.IsChecked = gRefreshSent
        MenuViewAll.IsChecked = gRefreshOtherFolders

        If gViewOtherFolders OrElse gViewInbox OrElse gViewSent Then
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

        ' Disable the listview context menu while a refresh is in progress to prevent actions
        Try
            If Me.ListView1 IsNot Nothing AndAlso Me.ListView1.ContextMenu IsNot Nothing Then
                Me.ListView1.ContextMenu.IsEnabled = False
            End If
        Catch
            ' silent failure consistent with existing style
        End Try

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
            If gUndoLogIndex > 0 Then MenuOptionEnabled("Undo", True)
            MenuOptionEnabled("View", True)
            MenuOptionEnabled("ToggleRead", True)

            Me.ListView1.Focus()

        Else

            If gRefreshOtherFolders OrElse gRefreshSent OrElse gRefreshInbox Then
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

        ' Re-enable the listview context menu now that refresh is complete
        Try
            If Me.ListView1 IsNot Nothing AndAlso Me.ListView1.ContextMenu IsNot Nothing Then
                Me.ListView1.ContextMenu.IsEnabled = True
            End If
        Catch
            ' silent failure consistent with existing style
        End Try

        EndRefreshCursor()

        gIsRefreshing = False

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

    Private Sub StartProgressTimer()
        SyncLock gProgressUpdateTimerLock
            Interlocked.Exchange(gProgressCounter, 0)
            If gProgressUpdateTimer IsNot Nothing Then
                gProgressUpdateTimer.Dispose()
                gProgressUpdateTimer = Nothing
            End If
            gProgressUpdateTimer = New System.Threading.Timer(AddressOf OnProgressTimerTick, Nothing, TimeSpan.Zero, TimeSpan.FromSeconds(1))
        End SyncLock
    End Sub

    Private Sub StopProgressTimer()
        SyncLock gProgressUpdateTimerLock
            If gProgressUpdateTimer IsNot Nothing Then
                gProgressUpdateTimer.Dispose()
                gProgressUpdateTimer = Nothing
            End If
            Interlocked.Exchange(gProgressCounter, 0)
        End SyncLock
    End Sub

    Private Sub OnProgressTimerTick(state As Object)
        Dim currentValue As Double = Interlocked.Read(gProgressCounter)
        Me.Dispatcher.BeginInvoke(Sub()
                                      Dim target = Math.Min(currentValue, Me.ProgressBar1.Maximum)
                                      SetProgressBarValue(target)
                                  End Sub)
    End Sub

    Public Class ListViewRowClass

        Public Enum ChainIndicatorValues As Integer
            NotPartOfAChain = 0
            TopOfTheChain = 1
            MiddleOfTheChain = 2
            EndOfTheChain = 3
        End Enum

        Private _MailBoxName As String
        Private _Index As Integer
        Private _FixedIndex As Integer
        Private _ChainIndicator As Integer
        Private _Subject As String
        Private _Trailer As String
        Private _From As String
        Private _xTo As String
        Private _DateTime As String
        Private _OriginalFolder As Integer
        Private _RecommendedFolderFinal As Integer
        Private _RecommendedFolder1 As Integer
        Private _RecommendedFolder2 As Integer
        Private _RecommendedFolder3 As Integer
        Private _OutlookEntryID As String
        Private _Unread As FontWeight

        Public Property MailBoxName() As String
            Get
                Return Me._MailBoxName
            End Get
            Set(ByVal value As String)
                Me._MailBoxName = value
            End Set
        End Property

        Public Property Index() As Integer
            Get
                Return Me._Index
            End Get
            Set(ByVal value As Integer)
                Me._Index = value
            End Set
        End Property

        Public Property FixedIndex() As Integer
            Get
                Return Me._FixedIndex
            End Get
            Set(ByVal value As Integer)
                Me._FixedIndex = value
            End Set
        End Property

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

        Public Property Subject() As String
            Get
                Return Me._Subject
            End Get
            Set(ByVal value As String)
                Me._Subject = value
            End Set
        End Property

        Public Property Trailer() As String
            Get
                Return Me._Trailer
            End Get
            Set(ByVal value As String)
                Me._Trailer = value
            End Set
        End Property

        Public Property From() As String
            Get
                Return Me._From
            End Get
            Set(ByVal value As String)
                Me._From = value
            End Set
        End Property

        Public Property xTo() As String
            Get
                Return Me._xTo
            End Get
            Set(ByVal value As String)
                Me._xTo = value
            End Set
        End Property

        Public Property DateTime() As Date
            Get
                Return Me._DateTime
            End Get
            Set(ByVal value As Date)
                Me._DateTime = value
            End Set
        End Property

        Public Property OriginalFolder() As Integer
            Get
                Return Me._OriginalFolder
            End Get
            Set(ByVal value As Integer)
                Me._OriginalFolder = value
            End Set
        End Property

        Public Property RecommendedFolderFinal() As Integer
            Get
                Return Me._RecommendedFolderFinal
            End Get
            Set(ByVal value As Integer)
                Me._RecommendedFolderFinal = value
            End Set
        End Property

        Public Property RecommendedFolder1() As Integer
            Get
                Return Me._RecommendedFolder1
            End Get
            Set(ByVal value As Integer)
                Me._RecommendedFolder1 = value
            End Set
        End Property

        Public Property RecommendedFolder2() As Integer
            Get
                Return Me._RecommendedFolder2
            End Get
            Set(ByVal value As Integer)
                Me._RecommendedFolder2 = value
            End Set
        End Property

        Public Property RecommendedFolder3() As Integer
            Get
                Return Me._RecommendedFolder3
            End Get
            Set(ByVal value As Integer)
                Me._RecommendedFolder3 = value
            End Set
        End Property

        Public Property OutlookEntryID() As String
            Get
                Return Me._OutlookEntryID
            End Get
            Set(ByVal value As String)
                Me._OutlookEntryID = value
            End Set
        End Property

        Public Property UnRead() As FontWeight
            Get
                Return Me._Unread
            End Get
            Set(ByVal value As FontWeight)
                Me._Unread = value
            End Set
        End Property

        Public Function Clone() As ListViewRowClass

            Dim copy As New ListViewRowClass

            copy.MailBoxName = Me.MailBoxName
            copy.Index = Me.Index
            copy.FixedIndex = Me.FixedIndex
            copy.ChainIndicator = Me.ChainIndicator

            copy.Subject = Me.Subject
            copy.Trailer = Me.Trailer
            copy.From = Me.From
            copy.xTo = Me.xTo
            copy.DateTime = Me.DateTime

            copy.OriginalFolder = Me.OriginalFolder
            copy.RecommendedFolder1 = Me.RecommendedFolder1
            copy.RecommendedFolder2 = Me.RecommendedFolder2
            copy.RecommendedFolder3 = Me.RecommendedFolder3
            copy.RecommendedFolderFinal = Me.RecommendedFolderFinal

            copy.OutlookEntryID = Me.OutlookEntryID
            copy.UnRead = Me.UnRead

            Return copy

        End Function

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
            MenuOptionEnabled("Refresh", (gRefreshInbox OrElse gRefreshSent OrElse gRefreshOtherFolders))
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

            If (gCurrentSortOrder = "Mailbox") OrElse (gCurrentSortOrder = "Subject") Then
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

            Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

            Call ShowMessageBox("FileFriendly - Set ListView Item Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 currentMethodName & " - " & ex.ToString,
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

            Dim cur = FinalRecommendationTable(x)
            If cur.Trailer.Length = 0 Then Continue For

            Dim prev = FinalRecommendationTable(x - 1)
            If prev Is Nothing OrElse cur Is Nothing Then Continue For

            If (prev.Trailer = cur.Trailer) Then
                'If prev.Subject = cur.Subject AndAlso prev.Trailer = cur.Trailer Then
                cur.ChainIndicator = ListViewRowClass.ChainIndicatorValues.MiddleOfTheChain
            End If

        Next

        'Set tops
        For x As Integer = 0 To FinalRecommendationTable.Length - 2

            Dim cur = FinalRecommendationTable(x)
            If cur.Trailer.Length = 0 Then Continue For

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
            If cur.Trailer.Length = 0 Then Continue For

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

            If last IsNot Nothing AndAlso last.ChainIndicator = ListViewRowClass.ChainIndicatorValues.MiddleOfTheChain Then
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

            If Not (gRefreshInbox OrElse gRefreshSent OrElse gRefreshOtherFolders) Then
                Me.lblMainMessageLine.Content = "0 e-mails"
                Exit Try
            End If

            Dim NewRecommendationTable(gFinalRecommendationTable.Length - 1) As ListViewRowClass

            For x As Integer = 0 To gFinalRecommendationTable.Length - 1

                If gFinalRecommendationTable(x) Is Nothing Then
                    Continue For
                End If

                If gFinalRecommendationTable(x).Index = -1 Then
                    Continue For
                End If

                Dim row = gFinalRecommendationTable(x)
                If row Is Nothing Then Continue For

                If gHiddenEntryIDs.Contains(row.OutlookEntryID) Then
                    ' This item is hidden, skip it
                    Continue For
                End If

                MessageWasRead = (row.UnRead = System.Windows.FontWeights.Normal)

                If (gViewRead AndAlso MessageWasRead) OrElse (gViewUnRead AndAlso (Not MessageWasRead)) Then

                    InboxItem = gFolderTable(row.OriginalFolder).FolderType = FolderTableType.Inbox
                    SentItem = gFolderTable(row.OriginalFolder).FolderType = FolderTableType.SentItems
                    NeitherInboxNorSentItem = Not (InboxItem OrElse SentItem)

                    If (gViewInbox AndAlso InboxItem) OrElse
                       (gViewSent AndAlso SentItem) OrElse
                       (gViewOtherFolders AndAlso NeitherInboxNorSentItem) Then

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

            If (gCurrentSortOrder = "Mailbox") OrElse (gCurrentSortOrder = "Subject") Then
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

#If DEBUG Then
            Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name
            Console.WriteLine(currentMethodName & " - " & ex.ToString)
#End If

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

        ' Remove arrow _from previously sorted header
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

            ' Remove arrow _from previously sorted header
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

    Private Sub RefreshBackGroundTask(ByVal InitialLoad As Boolean, ByVal MSOutlookDrivenEvent As Boolean, ByVal QuickRefresh As Boolean)


#If DEBUG Then

        'time how long the overall process takes (when in debug mode)
        Dim swOverall As Stopwatch = Stopwatch.StartNew()
        swOverall.Start()

#End If

        Try

            StartProgressTimer()

            SetMousePointer(Cursors.Wait)

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

                    Collection_of_folders_to_exclude = My.Settings.ExcludedScanFolders 'list of all folders to be excluded _from scan

                    Collection_of_folders_to_exclude_is_empty = (Collection_of_folders_to_exclude Is Nothing)

                    Me.Dispatcher.BeginInvoke(New ShowFoldersCallback(AddressOf ShowFolders), New Object() {})

                    FindAllFolders()

                    If gCancelRefresh Then GoTo CleanExit

                End If

            End If


            gMinimizeMaximizeAllowed = True

            Me.Dispatcher.BeginInvoke(New ShowFoldersCallback(AddressOf ShowFolders), New Object() {})

            If gRefreshInbox OrElse gRefreshSent OrElse gRefreshOtherFolders Then

                If lTotalEMailsToBeReviewed > 0 OrElse (MSOutlookDrivenEvent AndAlso gEmailTableIndex > 0) Then

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

            Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

            Call ShowMessageBox("FileFriendly - Refresh Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 currentMethodName & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        Finally

            UpdateSortHeaderGlyph()
            StopProgressTimer()
            SetMousePointer(Cursors.Arrow)
            MemoryManagement.FlushMemory()
            gIsRefreshing = False

        End Try

#If DEBUG Then

        swOverall.Stop()
        Dim ts As TimeSpan = TimeSpan.FromMilliseconds(swOverall.ElapsedMilliseconds)
        Console.WriteLine($"Overall time to refresh: {ts.Hours} hours, {ts.Minutes} minutes, {ts.Seconds} seconds")
        Console.WriteLine("Total e-mails reviewed: " & lTotalEMailsToBeReviewed.ToString)
        Console.WriteLine("E-mails / second: " & (lTotalEMailsToBeReviewed / (swOverall.ElapsedMilliseconds / 1000)).ToString("F2"))
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

            If gCancelRefresh Then GoTo EarlyExit

            ' sw.Stop()
            ' Console.WriteLine(sw.ElapsedMilliseconds.ToString) : sw.Stop()

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

            ' First xStep: locate Inbox/Sent and best delete folder (Deleted Items / Deleted / Trash) per store
            For x As Integer = 0 To gFolderTable.Length - 1

                gFolderTable(x).FolderType = FolderTableType.OtherFolders

                Dim fInfo As FolderInfo = gFolderTable(x)
                If fInfo.DefaultItemType <> Microsoft.Office.Interop.Outlook.OlItemType.olMailItem Then
                    Continue For
                End If

                Dim nameUpper As String = System.IO.Path.GetFileName(fInfo.FolderPath).Trim().ToUpperInvariant()

                ' Track Inbox folders globally and per mailbox
                If nameUpper = "INBOX" Then
                    gFolderTable(x).FolderType = FolderTableType.Inbox
                    Continue For
                End If

                ' Track Sent folders globally and per mailbox
                Select Case nameUpper
                    Case "SENT", "SENT ITEMS", "SENT MAIL"
                        gFolderTable(x).FolderType = FolderTableType.SentItems
                        Continue For
                End Select

                gFolderTable(x).FolderType = FolderTableType.OtherFolders

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
            ' Console.WriteLine("Inboxes: " & gInboxFolderIndices.NumberOfSelectedItems & " Sent: " & gSentFolderIndices.NumberOfSelectedItems)
#End If

            gFolderButtonsOnOptionsWindowEnabled = True
            gOptionsWindow?.SafelyEnableOptionsFolderButtons()

            Dim ToolTipMessage As String = ""
            Dim ProgressBarMaxValue As Double

            If gRefreshOtherFolders OrElse gRefreshInbox OrElse gRefreshSent Then

                If gRefreshOtherFolders Then

                    ToolTipMessage = "E-mails _from all included folders are being reviewed"

                    'ProcessBarMaxValue = 
                    ' 10 times the TotalEMails To Be Reviewed for processing all info but the workingBody +
                    ' 1 times the TotalEMails To Be Reviewed for processing the workingBody + 
                    ' a time factor doing the recommendations
                    'ProgressBarMaxValue = (3 * lTotalEMailsToBeReviewed) + Int(lTotalEMailsToBeReviewed * (1 + My.Settings.RatioOfRecommendationToProcessingTime + 0.01))
                    ProgressBarMaxValue = lTotalEMailsToBeReviewed

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

                ProgressBarMaxValue = 0

            End If

            Me.Dispatcher.BeginInvoke(New SetToolTipCallback(AddressOf SetToolTip), New Object() {ToolTipMessage})
            Me.Dispatcher.BeginInvoke(New SetProgressBarMaxValueCallback(AddressOf SetProgressBarMaxValue), New Object() {ProgressBarMaxValue})

EarlyExit:
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
        Dim isSentFolder As Boolean = String.Equals(folderName, "Sent", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(folderName, "Sent Items", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(folderName, "Sent Mail", StringComparison.OrdinalIgnoreCase)

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

    Private Function CreateTrailer(ByVal _dateAndTime As Date, ByVal _subject As String, _body As String) As String

        ' the trailer field is used to link emails together in chains

        ' the trailer is derived, where possible, from the email details and returned as a string representation of hashed value
        ' Outlooks ConversationID is not used as it is not fully reliable

        ' when creating the trailer the email's body is examined first and used if possible to create the trailer based on 
        ' the original email's sent date and time line and the original email's body

        ' the time stamp of the originating email is used in the trailer to help ensure its uniqueness
        ' ideally we would like to use "yyyy-MM-dd HH:mm:ss" for setting the trailer, however seconds are usually not reported in the email's sent date and time line 
        ' beyond that some system round the seconds to the nearest minute, while others truncate the seconds, this makes minutes unreliable for trailer creation, and hours
        ' (albeit less likely) unreliable too
        ' so this makes it problematic to use date and time at all, but its needed as using the subject alone can lead to too many different emails having the same trailer
        ' therefore we use "yyyy-MM-dd" to ensure consistent trailer creation with errors for chain id potentially happening only if the originating email was sent at midnight +/-30 seconds

        Const DatePersonFormat As String = "yyyy-MM-dd"

        Dim trailer As String

        If String.IsNullOrEmpty(_body.Trim()) Then

            If String.IsNullOrEmpty(_subject.Trim()) Then

                ' set trailer to a random number between 1 and 10,000,000 to all but completely ensure different emails with no subject or body have different trailers
                Dim random As New Random()
                trailer = random.Next(1, 10000001).ToString()

            Else

                ' the email has a subject but no body - create the trailer on that basis
                Dim originatingDateAndTimeString As String = _dateAndTime.ToString(DatePersonFormat, System.Globalization.CultureInfo.InvariantCulture)
                trailer = originatingDateAndTimeString & _subject.Trim()

            End If

        Else

            ' the originatingDateAndTimeString will be set to the originating email's date and time string in the format "yyyy-MM-dd HH:mm"
            ' note seconds are not included in the format as they are not always reported in the email's sent date and time line
            Dim originatingDateAndTimeString As String = String.Empty

            Dim workingBody = _body.Trim()

            Dim subjectIndex = workingBody.LastIndexOf("Subject:", StringComparison.OrdinalIgnoreCase)

            Dim sentIndex As Integer = -1

            ' find the originating date and time 

            If subjectIndex = -1 Then

                ' there was no subject line in the body of the email, which means this is an originating email; create the trailer on that basis

                originatingDateAndTimeString = _dateAndTime.ToString(DatePersonFormat, System.Globalization.CultureInfo.InvariantCulture)

            Else

                ' there was a subject line in the body of the email, which means this is a subsequent email in a chain; create the trailer on that basis

                ' find the first sent date and time which is before the last subject line in the body of the email - this will be the originating email's sent date and time

                sentIndex = workingBody.LastIndexOf("Sent:", subjectIndex, subjectIndex + 1, StringComparison.OrdinalIgnoreCase)

                If sentIndex > -1 Then

                    sentIndex += 6 ' length of "Sent: "

                    originatingDateAndTimeString = workingBody.Substring(sentIndex)

                    ' truncate anything after the first line break in originatingDateAndTimeString
                    Dim lineBreakIndex As Integer = originatingDateAndTimeString.IndexOfAny(New Char() {Chr(10), Chr(13)})
                    If lineBreakIndex > -1 Then
                        originatingDateAndTimeString = originatingDateAndTimeString.Substring(0, lineBreakIndex).Trim()
                    End If

                    ' ensue originatingDateAndTimeString is in this format: "yyyy-MM-dd HH:mm"
                    Dim parsedDate As Date
                    Dim workindDateAndTime As Date
                    If Date.TryParse(originatingDateAndTimeString, parsedDate) Then
                        workindDateAndTime = parsedDate
                    Else
                        workindDateAndTime = _dateAndTime
                    End If

                    originatingDateAndTimeString = workindDateAndTime.ToString(DatePersonFormat, System.Globalization.CultureInfo.InvariantCulture)

                    ' set the workingBody to the originating email's body - that is everything after the last subject line
                    workingBody = workingBody.Remove(0, subjectIndex)
                    lineBreakIndex = workingBody.IndexOfAny(New Char() {Chr(10), Chr(13)})
                    If lineBreakIndex > -1 Then
                        workingBody = workingBody.Substring(lineBreakIndex).Trim()
                    End If

                End If

            End If

            trailer = originatingDateAndTimeString & workingBody


        End If

        trailer = Regex.Replace(trailer, "\s+", "") ' remove all whitespace characters to ensure consistent hashing

        'Console.Write(_subject & "    " & trailer & " ")

        ' compute a MD5-based fingerprint for the trailer and return it as a hexadecimal string

        Dim hashBytes As Byte() = md5Obj.ComputeHash(System.Text.Encoding.ASCII.GetBytes(If(trailer, String.Empty)))

        trailer = BitConverter.ToString(hashBytes).Replace("-", "")

#If DEBUG Then
        ' System.IO.File.AppendAllText("c:\temp\test.txt", trailer & " , """ & _subject & """" & Environment.NewLine)
#End If

        Return trailer

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
        strCollection = My.Settings.ExcludedScanFolders 'list of all folders to be excluded _from scan

        With lBlankEMailDetailRecord

            .sSubject = ""
            .sTrailer = ""
            .sTo = ""
            .sFrom = ""
            .sDateAndTime = Now
            .sOutlookEntryID = ""
            .sUnRead = System.Windows.FontWeights.Bold

        End With

        '***************************************************************************
        'Step 2 add all info  
        '***************************************************************************

        ' below we will add all other folders first, and inbox and sent items second
        ' this is done as the add of the other folders can be long running, while the inbox and sent items are quick
        ' in this way any emails sent or received during the long running other folder process will be picked up in the inbox and sent item processing that comes after

        Dim PopulateFoldersUsingPreviousData As Boolean = True

        Dim CurrentMailBoxName As String = ""

        '***************************************************************************
        'Step 2A add other folders
        '***************************************************************************

        If gRefreshOtherFolders AndAlso (Not MSOutlookDrivenEvent) AndAlso (Not QuickRefresh) Then

            For x As Integer = 0 To gFolderTableIndex

                If gCancelRefresh Then Exit For

                If gFolderTable(x).FolderType = FolderTableType.OtherFolders Then

                    If IsAMailboxBoxPath(gFolderTable(x).FolderPath) Then
                        CurrentMailBoxName = GetMailboxNameFromFolderPath(gFolderTable(x).FolderPath, gFolderTable(x).StoreID)
                    End If

                    If (strCollection.IndexOf(gFolderNamesTable(x)) = -1) OrElse Collection_of_folders_to_exclude_is_empty Then

                        If gFolderTable(x).DefaultItemType = Microsoft.Office.Interop.Outlook.OlItemType.olMailItem Then
                            Dim folder As Microsoft.Office.Interop.Outlook.MAPIFolder = oNS.GetFolderFromID(gFolderTable(x).EntryID, gFolderTable(x).StoreID)
                            Try
                                ProcessAllMailItemsInAFolder(x, folder, CurrentMailBoxName)
                            Finally
                                If folder IsNot Nothing Then
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(folder)
                                    folder = Nothing
                                End If
                            End Try
                        End If

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

        '***************************************************************************
        'Step 2B add inbox and sent 
        '***************************************************************************

        If gRefreshInbox OrElse gRefreshSent Then

            For x As Integer = 0 To gFolderTableIndex

                If gCancelRefresh Then Exit For

                If IsAMailboxBoxPath(gFolderTable(x).FolderPath) Then
                    CurrentMailBoxName = GetMailboxNameFromFolderPath(gFolderTable(x).FolderPath, gFolderTable(x).StoreID)
                End If

                If (gRefreshSent AndAlso (gFolderTable(x).FolderType = FolderTableType.SentItems)) OrElse
                   (gRefreshInbox AndAlso (gFolderTable(x).FolderType = FolderTableType.Inbox)) Then

                    If gFolderTable(x).DefaultItemType = Microsoft.Office.Interop.Outlook.OlItemType.olMailItem Then
                        Dim folder As Microsoft.Office.Interop.Outlook.MAPIFolder = oNS.GetFolderFromID(gFolderTable(x).EntryID, gFolderTable(x).StoreID)
                        Try
                            ProcessAllMailItemsInAFolder(x, folder, CurrentMailBoxName)
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

EarlyExit:

        ' sw.Stop()
        ' Console.WriteLine(sw.ElapsedMilliseconds)

    End Sub

    Private Sub ProcessAllMailItemsInAFolder(ByVal originalFolder As Integer,
                             ByVal folder As Microsoft.Office.Interop.Outlook.MAPIFolder,
                             ByVal mailboxName As String)

        If gCancelRefresh Then Exit Sub

        Dim items As Microsoft.Office.Interop.Outlook.Items = Nothing

        Try
            items = folder.Items
        Catch
            Exit Sub
        End Try

        ' Console.WriteLine("Processing folder: " & gFolderTable(originalFolder).FolderPath)

        Try

            Try
                Dim sortField As String = If(lWhenSent, "[SentOn]", "[ReceivedTime]")
                items.Sort(sortField, True)
            Catch
            End Try

            Dim itemCount As Integer = items.Count ' set the number of items in the folder as a variable (_to avoid having to access it repeatably in the line below)
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

                    ' Resolve a friendly "From" address (gets around a quirk in Outlook / Exchange for messages coming _from Exchange or certain connected accounts)
                    Dim exUser As Microsoft.Office.Interop.Outlook.ExchangeUser = Nothing

                    Try
                        If mail.Sender IsNot Nothing Then
                            If String.Equals(mail.SenderEmailType, "SMTP", StringComparison.OrdinalIgnoreCase) Then
                            Else
                                exUser = TryCast(mail.Sender.GetExchangeUser(), Microsoft.Office.Interop.Outlook.ExchangeUser)

                                If exUser IsNot Nothing AndAlso Not String.IsNullOrEmpty(exUser.PrimarySmtpAddress) Then
                                    friendlyFrom = exUser.PrimarySmtpAddress

                                End If
                            End If

                        End If
                    Catch
                    Finally
                        If exUser IsNot Nothing Then
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(exUser)
                            exUser = Nothing
                        End If
                    End Try

                    If gHiddenEntryIDs.Contains(mail.EntryID) Then
                        ' This item is hidden, skip it
                        ' Console.WriteLine("Skipping hidden e-mail: " & mail.Subject)
                    Else

                        With emailDetail

                            .sOriginalFolderReferenceNumber = originalFolder
                            .sOutlookEntryID = mail.EntryID
                            .sSubject = CleanUpSubjectLine(If(mail.Subject, String.Empty))
                            .sTo = If(mail.To, String.Empty)
                            .sFrom = If(friendlyFrom, String.Empty)
                            .sDateAndTime = If(lWhenSent, mail.SentOn, mail.ReceivedTime)
                            .sUnRead = If(mail.UnRead, System.Windows.FontWeights.Bold, System.Windows.FontWeights.Normal)
                            .sMailBoxName = mailboxName
                            .sTrailer = CreateTrailer(.sDateAndTime, .sSubject, If(mail.Body, String.Empty))

                        End With

                        gEmailTable(gEmailTableIndex) = emailDetail
                        gEmailTableIndex += 1

                    End If

                    Interlocked.Increment(gProgressCounter)

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

#End Region

#Region "Establish and set rankings"

    Private Sub EstablishRecommendations()

        'Dim sw As New Stopwatch
        'sw.Start()
        Try

            ' A second sort of the email table is required to subjects in order with their trailers
            ' this is because we want to group emails in the same chain together for scoring and establishing recommendations
            ' (x.e. one set of recommendations for all emails in the same chain)

            Array.Sort(gEmailTable, 0, gEmailTableIndex, EMailTableSorter.SubjectThenDateAsc)

            EstablishRatings_NumberOfUniqueEmailsInAFolder()

            EstablishRatings_Scoring()

        Catch ex As Exception

            Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

            Call ShowMessageBox("FileFriendly - Establish Recommendations Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 currentMethodName & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try

        ' sw.Stop()
        ' Console.WriteLine("EstablishRecommendations time: " & sw.ElapsedMilliseconds)

    End Sub

    Private Sub EstablishRatings_NumberOfUniqueEmailsInAFolder()

        'Set up for rating number of e-mails related to the same chain within a folder

        Try
            UniqueSubjectsMap.Clear()

            For x As Integer = 0 To gEmailTable.Length - 1
                Dim subjectAndTrailer As String = gEmailTable(x).sSubject & gEmailTable(x).sTrailer
                Dim folderRef As Integer = gEmailTable(x).sOriginalFolderReferenceNumber

                Dim folderCounts As Dictionary(Of Integer, Integer) = Nothing
                If Not UniqueSubjectsMap.TryGetValue(subjectAndTrailer, folderCounts) Then
                    folderCounts = New Dictionary(Of Integer, Integer)()
                    UniqueSubjectsMap(subjectAndTrailer) = folderCounts
                End If

                Dim count As Integer
                If folderCounts.TryGetValue(folderRef, count) Then
                    folderCounts(folderRef) = count + 1
                Else
                    folderCounts(folderRef) = 1
                End If
            Next

        Catch ex As Exception

            Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

            Call ShowMessageBox("FileFriendly - Establish Ratings Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 currentMethodName & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try

    End Sub

    Private Sub EstablishRatings_Scoring()

        'For each unique e-mail chain, rate the best folder to put it in

        '   1 point to each folder (excluding inbox and sent folders) that has an e-mail in it which belongs to the chain

        Try

            Dim CurrentSubjectAndTrailer As String = "|*| something unique |*|" & Chr(255)
            Dim PrevSubjectAndTrailer As String = ""

            Dim FinalScoringTable(gFolderTable.Length - 1) As Integer

            For i As Integer = 0 To gEmailTable.Length - 1

                PrevSubjectAndTrailer = CurrentSubjectAndTrailer
                CurrentSubjectAndTrailer = gEmailTable(i).sSubject & gEmailTable(i).sTrailer

                If CurrentSubjectAndTrailer = PrevSubjectAndTrailer Then

                    gEmailTable(i).sRecommendedFolder1ReferenceNumber = gEmailTable(i - 1).sRecommendedFolder1ReferenceNumber
                    gEmailTable(i).sRecommendedFolder2ReferenceNumber = gEmailTable(i - 1).sRecommendedFolder2ReferenceNumber
                    gEmailTable(i).sRecommendedFolder3ReferenceNumber = gEmailTable(i - 1).sRecommendedFolder3ReferenceNumber
                    gEmailTable(i).sRecommendedFolderFinalReferenceNumber = gEmailTable(i - 1).sRecommendedFolderFinalReferenceNumber

                Else

                    Dim folderCounts As Dictionary(Of Integer, Integer) = Nothing
                    UniqueSubjectsMap.TryGetValue(CurrentSubjectAndTrailer, folderCounts)

                    Array.Clear(FinalScoringTable, 0, FinalScoringTable.Length)

                    If folderCounts IsNot Nothing Then
                        For Each kvp As KeyValuePair(Of Integer, Integer) In folderCounts
                            If kvp.Key >= 0 AndAlso kvp.Key < FinalScoringTable.Length Then
                                FinalScoringTable(kvp.Key) = kvp.Value
                            End If
                        Next
                    End If

                    FindTheFolderWithTheGreatestScore(gEmailTable(i).sRecommendedFolder1ReferenceNumber, FinalScoringTable)
                    FindTheFolderWithTheGreatestScore(gEmailTable(i).sRecommendedFolder2ReferenceNumber, FinalScoringTable)
                    FindTheFolderWithTheGreatestScore(gEmailTable(i).sRecommendedFolder3ReferenceNumber, FinalScoringTable)
                    gEmailTable(i).sRecommendedFolderFinalReferenceNumber = gEmailTable(i).sRecommendedFolder1ReferenceNumber

                End If

            Next

            UniqueSubjectsMap.Clear()

        Catch ex As Exception

            Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

            Call ShowMessageBox("FileFriendly - Establish Ratings Scoring Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 currentMethodName & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try

    End Sub

    Private Sub FindTheFolderWithTheGreatestScore(ByRef ReferenceNumber As Integer, ByRef FinalScoringTable() As Integer)

        ' note gFolderTable and gFinalScoringTable are the same size
        ' and each index in gFolderTable corresponds to the same index in gFinalScoringTable

        Dim max As Integer = 0
        Dim MaxIndex As Integer = 0
        For x As Integer = 0 To FinalScoringTable.Length - 1
            ' find the score maximum for only other folders (excluding inbox or sent items - as we don't want to recommend filing an e-mail in them)
            If (FinalScoringTable(x) > max) AndAlso (gFolderTable(x).FolderType = FolderTableType.OtherFolders) Then
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
            'For index As Integer = 0 To index - 1
            '    Console.WriteLine(index.ToString & " " & gEmailTable(index).sSubject)
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


                    If gFolderTable(gEmailTable(lNextIndex).sOriginalFolderReferenceNumber).FolderType = FolderTableType.OtherFolders Then

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

                    ElseIf gFolderTable(gEmailTable(lNextIndex).sOriginalFolderReferenceNumber).FolderType = FolderTableType.Inbox Then

                        If gRefreshInbox Then lFlagThisEmailChain = True

                    Else

                        If gRefreshSent Then lFlagThisEmailChain = True

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

                            If (gFolderTable(gEmailTable(lNextIndex).sOriginalFolderReferenceNumber).FolderType = FolderTableType.OtherFolders) Then

                                lRecommendedIndexForAllEntriesInChainFinal = gEmailTable(lNextIndex).sRecommendedFolderFinalReferenceNumber
                                lRecommendedIndexForAllEntriesInChain1 = gEmailTable(lNextIndex).sRecommendedFolder1ReferenceNumber
                                lRecommendedIndexForAllEntriesInChain2 = gEmailTable(lNextIndex).sRecommendedFolder2ReferenceNumber
                                lRecommendedIndexForAllEntriesInChain3 = gEmailTable(lNextIndex).sRecommendedFolder3ReferenceNumber
                                lFlagThisEmailChain = True
                                Exit While

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

            Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

            Call ShowMessageBox("FileFriendly - Update List View Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 currentMethodName & " - " & ex.ToString,
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

        ' When a refresh is running, ignore Delete and Ctrl+Z and play a beep
        If gIsRefreshing Then
            Try
                Dim ctrlPressed As Boolean = (System.Windows.Input.Keyboard.Modifiers And System.Windows.Input.ModifierKeys.Control) = System.Windows.Input.ModifierKeys.Control
                If e.Key = System.Windows.Input.Key.Delete OrElse (e.Key = System.Windows.Input.Key.Z AndAlso ctrlPressed) Then
                    Try
                        Beep()
                    Catch
                        ' swallow any errors _from Beep to keep behaviour consistent
                    End Try
                    e.Handled = True
                    Return
                End If
            Catch
                ' keep silent on unexpected errors
            End Try
        End If

        ProcessKeyDown(e)
    End Sub

    Private Sub ListView1_ContextMenuOpening(ByVal sender As Object,
                                         ByVal e As ContextMenuEventArgs) _
                                         Handles ListView1.ContextMenuOpening

        ' If a refresh is underway, suppress the context menu _from opening.
        ' This prevents the user _from invoking actions while the underlying data is changing.
        If gIsRefreshing Then
            e.Handled = True
            Return
        End If

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


    <System.Diagnostics.DebuggerStepThrough()>
    Private Function BuildChainKey(ByVal row As ListViewRowClass) As String

        If row Is Nothing Then Return ""

        Return row.Trailer

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


#If DEBUG Then

    Private Sub PostRestoreSelectionDebugPrints()

        Static xStep As Integer = 1

        Dim www = Me.Dispatcher.BeginInvoke(Sub()
                                                Console.WriteLine("     ")

                                                Dim itemIndex As Integer = 0

                                                For Each obj In ListView1.Items
                                                    Dim row As ListViewRowClass = TryCast(obj, ListViewRowClass)
                                                    If row Is Nothing Then Continue For

                                                    Dim selectedText As String = "(no list items yet)"

                                                    Try
                                                        If ListView1 IsNot Nothing AndAlso ListView1.Items IsNot Nothing Then

                                                            If ListView1.SelectedItems.Contains(obj) Then
                                                                selectedText = "**** SELECTED ****"
                                                            Else
                                                                selectedText = "-- Not Selected --"
                                                            End If

                                                        End If
                                                    Catch
                                                        selectedText = "(selection unknown)"
                                                    End Try

                                                    Console.WriteLine("Step (" & xStep & ") row (" & itemIndex & ") " & vbTab & selectedText & vbTab & row.Subject & " " & row.OutlookEntryID & " " & row.Trailer)

                                                    itemIndex += 1

                                                Next
                                                xStep += 1
                                            End Sub)

        www.Wait()

    End Sub

#End If

    Delegate Sub RestoreSelectionCallback()
    Private Sub RestorePendingSelection()

        ' Prevent redundant re-entry only when a prior restore already produced a selection; allow a second xStep if nothing was selected yet.
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

            Dim TargetIndex As Integer


            If reason = SelectionRestoreReason.UserDelete OrElse reason = SelectionRestoreReason.OutlookDelete OrElse reason = SelectionRestoreReason.Hide Then

                ' select the next item below the last item which was deleted

                TargetIndex = snapshot.Entries.Last.Index


                If reason = SelectionRestoreReason.Hide Then

                    ' the count of selected items is subtracted for hidden items because unlike deleted or filed items those items were removed from the listview after the snapshot was taken
                    TargetIndex -= snapshot.Entries.Count

                End If


                If TargetIndex = (ListView1.Items.Count - 1) Then

                    ' we are at the bottom of the list so try to find a previous item that has not been deleted
                    ' so as to position the current selection at the end of the list once again

                    For i As Integer = TargetIndex To 0 Step -1
                        Dim row = TryCast(ListView1.Items(i), ListViewRowClass)
                        If row Is Nothing Then Continue For
                        If String.IsNullOrEmpty(row.OutlookEntryID) Then Continue For
                        If row.Index = -1 Then Continue For
                        TargetIndex = i
                        gCurrentlySelectedListViewItemIndex = i
                        ListView1.SelectedIndex = i
                        Exit For
                    Next

                Else

                    ' we are somewhere above the bottom of the list so just select the next item

                    If TargetIndex > ListView1.Items.Count - 1 Then TargetIndex = ListView1.Items.Count - 1

                    If TargetIndex >= 0 AndAlso TargetIndex < ListView1.Items.Count Then
                        ListView1.SelectedIndex = TargetIndex
                        gCurrentlySelectedListViewItemIndex = TargetIndex
                    Else
                        gCurrentlySelectedListViewItemIndex = 0
                        ListView1.SelectedIndex = -1
                    End If

                End If

                ' reselect chain members if enabled.

                If ListView1.SelectedIndex = -1 Then
                Else
                    If gAutoChainSelect Then
                        Dim anchorRow = TryCast(ListView1.SelectedItem, ListViewRowClass)
                        If anchorRow IsNot Nothing Then
                            Dim chainKey As String = BuildChainKey(anchorRow)
                            If Not String.IsNullOrEmpty(chainKey) Then
                                For i As Integer = 0 To ListView1.Items.Count - 1
                                    Dim row = TryCast(ListView1.Items(i), ListViewRowClass)
                                    If row Is Nothing Then Continue For
                                    If BuildChainKey(row) = chainKey Then
                                        If Not ListView1.SelectedItems.Contains(ListView1.Items(i)) Then
                                            ListView1.SelectedItems.Add(ListView1.Items(i))
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If

                End If


                If ListView1.SelectedItems.Count > 0 Then
                    Dim firstSelectedIndex As Integer = -1
                    For i As Integer = 0 To ListView1.Items.Count - 1
                        If ListView1.SelectedItems.Contains(ListView1.Items(i)) Then
                            firstSelectedIndex = i
                            Exit For
                        End If
                    Next

                    If firstSelectedIndex >= 0 Then
                        ListView1.SelectedIndex = firstSelectedIndex
                        gCurrentlySelectedListViewItemIndex = firstSelectedIndex
                        gPendingSelectionApplied = True
                        ListView1.UpdateLayout()
                        Dim selectedItem = ListView1.SelectedItem
                        ListView1.ScrollIntoView(selectedItem)
                        Dispatcher.BeginInvoke(Sub()
                                                   ListView1.UpdateLayout()
                                                   Dispatcher.BeginInvoke(Sub()
                                                                              Dim selectedContainer = TryCast(ListView1.ItemContainerGenerator.ContainerFromIndex(ListView1.SelectedIndex), System.Windows.Controls.ListViewItem)
                                                                              If selectedContainer IsNot Nothing Then
                                                                                  selectedContainer.Focus()
                                                                                  System.Windows.Input.Keyboard.Focus(selectedContainer)
                                                                              End If
                                                                          End Sub, System.Windows.Threading.DispatcherPriority.ApplicationIdle)
                                               End Sub, System.Windows.Threading.DispatcherPriority.Loaded)
                    End If
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
                            If Not String.IsNullOrEmpty(chainKey) Then
                                For i As Integer = 0 To ListView1.Items.Count - 1
                                    Dim row = TryCast(ListView1.Items(i), ListViewRowClass)
                                    If row Is Nothing Then Continue For
                                    If BuildChainKey(row) = chainKey Then
                                        If Not ListView1.SelectedItems.Contains(ListView1.Items(i)) Then
                                            ListView1.SelectedItems.Add(ListView1.Items(i))
                                        End If
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

            'Track chain keys _from the snapshot to reapply chain selections when appropriate.
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

            'Skip EntryID matching for delete operations since deleted items are being removed, not preserved.
            If reason <> SelectionRestoreReason.UserDelete AndAlso reason <> SelectionRestoreReason.OutlookDelete Then
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
            End If

            'Ensure all surviving rows with matching Outlook IDs _from the snapshot are reselected, even if indices shifted.
            'Skip this for delete operations since we want to select the next available item/chain, not preserve deleted item IDs.
            If reason <> SelectionRestoreReason.UserDelete AndAlso reason <> SelectionRestoreReason.OutlookDelete Then
                If selectedIds.Count > 0 Then
                    For i As Integer = 0 To ListView1.Items.Count - 1
                        Dim row = TryCast(ListView1.Items(i), ListViewRowClass)
                        If row Is Nothing Then Continue For
                        If Not String.IsNullOrEmpty(row.OutlookEntryID) AndAlso selectedIds.Contains(row.OutlookEntryID) Then
                            targetIndices.Add(i)
                        End If
                    Next
                End If
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

                ' Guard against stale indices _from the snapshot; skip anything outside the current bounds before applying.
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
                Dim firstSelectedIndex As Integer = -1
                For i As Integer = 0 To ListView1.Items.Count - 1
                    If ListView1.SelectedItems.Contains(ListView1.Items(i)) Then
                        firstSelectedIndex = i
                        Exit For
                    End If
                Next

                If firstSelectedIndex < 0 Then
                    firstSelectedIndex = 0
                End If

                ListView1.SelectedIndex = firstSelectedIndex
                gCurrentlySelectedListViewItemIndex = firstSelectedIndex
                gPendingSelectionApplied = True

                ListView1.UpdateLayout()
                ListView1.ScrollIntoView(ListView1.Items(firstSelectedIndex))
                Dispatcher.BeginInvoke(Sub()
                                           ListView1.UpdateLayout()
                                           Dispatcher.BeginInvoke(Sub()
                                                                      Dim selectedContainer = TryCast(ListView1.ItemContainerGenerator.ContainerFromIndex(ListView1.SelectedIndex), System.Windows.Controls.ListViewItem)
                                                                      If selectedContainer IsNot Nothing Then
                                                                          selectedContainer.Focus()
                                                                          System.Windows.Input.Keyboard.Focus(selectedContainer)
                                                                      End If
                                                                  End Sub, System.Windows.Threading.DispatcherPriority.ApplicationIdle)
                                       End Sub, System.Windows.Threading.DispatcherPriority.Loaded)
            End If

            'Refresh detail view to reflect the restored selection.
            UpdateDetails()

        Finally
            'Reattach selection handler now that restore is complete.
            AddHandler ListView1.SelectionChanged, AddressOf ListView1_SelectionChanged

#If DEBUG Then
            ' PostRestoreSelectionDebugPrints()
#End If

        End Try

    End Sub


    Private Sub SelectAllMembersOfAnEmailChain()

        Try
            If ListView1.SelectedItems.Count = 0 Then Exit Sub

            Dim SelectedListViewItem = TryCast(ListView1.SelectedItems.Item(ListView1.SelectedItems.Count - 1), ListViewRowClass)
            If SelectedListViewItem Is Nothing Then Exit Sub

            Dim chainKey = BuildChainKey(SelectedListViewItem)
            For i = 0 To ListView1.Items.Count - 1
                Dim row = TryCast(ListView1.Items(i), ListViewRowClass)
                If row Is Nothing Then Continue For
                If BuildChainKey(row) = chainKey Then
                    If Not ListView1.SelectedItems.Contains(ListView1.Items(i)) Then
                        ListView1.SelectedItems.Add(ListView1.Items(i))
                    End If
                End If
            Next
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
    Private Function IsATrailerFoundInMultipleMailboxes() As Boolean

        Dim ReturnValue As Boolean

        If (gFinalRecommendationTable Is Nothing) OrElse (gFinalRecommendationTable.Length = 0) OrElse (_TotalMailBoxes < 2) Then

            ReturnValue = False

        Else

            Dim trailerToMailboxes As New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)

            For Each row As ListViewRowClass In gFinalRecommendationTable

                If row Is Nothing Then Continue For
                If row.Index = -1 Then Continue For

                Dim trailer As String = row.Trailer
                Dim mailbox As String = row.MailBoxName

                If String.IsNullOrEmpty(trailer) OrElse String.IsNullOrEmpty(mailbox) Then Continue For

                If Not trailerToMailboxes.ContainsKey(trailer) Then
                    trailerToMailboxes(trailer) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                End If

                trailerToMailboxes(trailer).Add(mailbox)

            Next

            For Each kvp In trailerToMailboxes
                If kvp.Value.Count > 1 Then
                    ReturnValue = True
                    Exit For
                End If
            Next

        End If

        Return ReturnValue

    End Function

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

            If (gCurrentSortOrder <> "Subject") AndAlso IsATrailerFoundInMultipleMailboxes() Then
                Me.lblMainMessageLine.Content &= " Note: some e-mail chains span multiple mailboxes, sort by Subject to view all e-mails chains grouped together"
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub UpdateMenuToogleReadUnread()
        Try
            Dim selected = TryCast(ListView1.SelectedItem, ListViewRowClass)
            If selected Is Nothing Then
                Exit Try
            End If

            If selected.UnRead = System.Windows.FontWeights.Bold Then
                Me.MenuToggle.Header = "Mark as unread"
            Else
                Me.MenuToggle.Header = "Mark as read"
            End If

            'Me.MenuToggle.Visibility = Windows.Visibility.Visible

        Catch ex As Exception
            Me.MenuToggle.Header = "Toggle Read/Unread"
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
            UpdateMenuToogleReadUnread()
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
    MenuOpen.Click, MenuHide.Click, MenuDelete.Click, MenuExit.Click, MenuToggle.Click,
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
                Dim store As Microsoft.Office.Interop.Outlook.Store = Nothing
                Try
                    store = oNS.GetStoreFromID(storeId)
                    If store IsNot Nothing AndAlso store.DisplayName IsNot Nothing Then
                        Return store.DisplayName
                    End If
                Catch
                    ' fall back to folderPath parsing below
                Finally
                    If store IsNot Nothing Then
                        Try
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(store)
                        Catch
                        End Try
                        store = Nothing
                    End If
                End Try
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
    Private Sub ActivateMenuNow()

        Me.Menu1.Focus()

    End Sub

    Private Sub PerformActionByProxyNow()

        PerformAction(gProxyAction)

    End Sub

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
                Me.MenuContextToggleRead.Header = "Mark as read"
            Else
                Me.MenuContextToggleRead.Header = "Mark as unread"
            End If

            Me.MenuContextToggleRead.Visibility = Windows.Visibility.Visible

        Catch ex As Exception
            ' Keep context‑menu failures silent, consistent with rest of file
            Me.MenuContextToggleRead.Header = "Toggle Read/Unread"
        End Try

    End Sub

    Private Sub StartRefresh(ByVal QuickRefresh As Boolean)

        gRefreshConfirmed = False

        Try
            If gIsRefreshing Then
                gCancelRefresh = True
                Exit Sub
            End If

            If gUndoLogIndex > 0 Then

                If ShowMessageBox("FileFriendly",
                       CustomDialog.CustomDialogIcons.Question,
                       "Please note",
                       "If you refresh you will no longer be able to undo the changes you have made up until now." & vbCrLf & vbCrLf &
                       "Would you still like to refresh?",
                       "You will however be able to undo future changes.",
                       "",
                       CustomDialog.CustomDialogIcons.None,
                       CustomDialog.CustomDialogButtons.YesNo,
                       CustomDialog.CustomDialogResults.No) = CustomDialog.CustomDialogResults.No Then
                    Exit Sub
                End If

            End If

            MenuOptionEnabled("Undo", False)
            gUndoLogIndex = 0

            If QuickRefresh Then
                gRefreshConfirmed = True
            Else
                gPickARefreshModeWindow = New PickARefreshMode
                gPickARefreshModeWindow.ShowDialog()
                gPickARefreshModeWindow = Nothing
            End If

            If gRefreshConfirmed Then

                If gRefreshInbox OrElse gRefreshSent OrElse gRefreshOtherFolders Then
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
                        OpenAnEmail()
                    End If

                Case Is = "File", "Delete", "Hide", "ToggleRead"

                    If ConfirmActionMessage(Action) Then
                        ActionRequestAgainstAllSelectedItems(Action, Me.ListView1)
                    End If

                Case Is = "Options"

                    gARefreshIsRequired = False

                    gOptionsWindow = New OptionsWindow
                    gOptionsWindow.ShowDialog()
                    gOptionsWindow = Nothing

                    gRefreshInbox = My.Settings.ScanInbox
                    gRefreshSent = My.Settings.ScanSent
                    gRefreshOtherFolders = My.Settings.ScanAll
                    gAutoChainSelect = My.Settings.AutoChainSelect

                    If gARefreshIsRequired Then
                        RefreshGrid(False, False, True)
                    End If

                Case Is = "Undo"

                    If My.Settings.ConfirmUndo Then
                        If ConfirmActionMessage(Action) Then RestoreFromUndoLog()
                    Else
                        RestoreFromUndoLog()
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

                Case Is = "ViewOtherFolders"
                    gViewOtherFolders = flag
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

            Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

            Call ShowMessageBox("FileFriendly - Perform Action Error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 currentMethodName & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        End Try

        Me.Cursor = Cursors.Arrow

    End Sub

    Private Sub ValidateInboxSentFoldersCombinatation()

        If gViewInbox OrElse gViewSent OrElse gViewOtherFolders OrElse (Me.MenuRefresh.Foreground Is gForegroundColourAlert) Then

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

        If (Me.MenuViewRead.Foreground Is gForegroundColourAlert) OrElse (Me.MenuViewInbox.Foreground Is gForegroundColourAlert) Then
            Me.MenuView.Foreground = gForegroundColourAlert
        Else
            Me.MenuView.Foreground = gForegroundColourEnabled
        End If

    End Sub

    Private Sub ValidateReadUnReadCombinatation()

        If gViewRead OrElse gViewUnRead OrElse (Me.MenuRefresh.Foreground Is gForegroundColourAlert) Then

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

        If (Me.MenuViewRead.Foreground Is gForegroundColourAlert) OrElse (Me.MenuViewInbox.Foreground Is gForegroundColourAlert) Then
            Me.MenuView.Foreground = gForegroundColourAlert
        Else
            Me.MenuView.Foreground = gForegroundColourEnabled
        End If

    End Sub

    Private Sub ShutDown()

        SyncLock gRefreshGateLock
            If gIsRefreshing Then
                gCancelRefresh = True
            End If
        End SyncLock

        ' Wait for refresh to finish, but with a timeout to avoid infinite loop
        Dim waitStart As DateTime = DateTime.Now
        Dim maxWait As TimeSpan = TimeSpan.FromSeconds(10) ' 10 seconds timeout

        While gIsRefreshing
            System.Threading.Thread.Sleep(100)
            If DateTime.Now - waitStart > maxWait Then
                Exit While
            End If
        End While

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

            Dim entryId As String = selectedRow.OutlookEntryID
            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = Nothing

            ' Retry GetItemFromID once if we see an RPC disconnect
            Dim attempt As Integer = 0
            While attempt < 2 AndAlso mailItem Is Nothing

                Try

                    mailItem = TryCast(oNS.GetItemFromID(entryId), Microsoft.Office.Interop.Outlook.MailItem)

                Catch comEx As System.Runtime.InteropServices.COMException

                    Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

                    Const RPC_E_SERVER_UNAVAILABLE As Integer = &H800706BA
                    Const RPC_E_DISCONNECTED As Integer = &H800706BE

                    If (comEx.HResult = RPC_E_SERVER_UNAVAILABLE OrElse comEx.HResult = RPC_E_DISCONNECTED) AndAlso attempt = 0 Then
                        ' Drop and rebuild the Outlook session once
                        oNS = Nothing
                        oApp = Nothing

                        ' Loop will retry with fresh session
                        mailItem = Nothing
                    Else

                        Call ShowMessageBox("FileFriendly - E-mail open failed",
                             CustomDialog.CustomDialogIcons.Stop,
                             "FileFriendly could Not open the selected e-mail",
                             "FileFriendly could Not open the selected e-mail in Outlook. (1)",
                             currentMethodName & " - " & comEx.ToString,
                             "",
                             CustomDialog.CustomDialogIcons.None,
                             CustomDialog.CustomDialogButtons.OK,
                             CustomDialog.CustomDialogResults.OK)

                        Exit Try
                    End If

                Catch ex As Exception

                    Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

                    Call ShowMessageBox("FileFriendly - E-mail open failed",
                         CustomDialog.CustomDialogIcons.Stop,
                         "FileFriendly could Not open the selected e-mail",
                         "FileFriendly could Not open the selected e-mail in Outlook. (2)",
                         currentMethodName & " - " & ex.ToString,
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

            ' if the email isn't already marked as read then mark it read in the listview
            Dim index As Integer = ListView1.SelectedIndex
            If index >= 0 AndAlso ListView1.Items(index).UnRead = System.Windows.FontWeights.Bold Then
                ToggleReadStateOfASelectedItemIn_TheListView(index)
            End If

            If mailItem IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                mailItem = Nothing
            End If

        Catch ex As Exception

            Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

            Call ShowMessageBox("FileFriendly - Open an e-mail error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error." & vbCrLf & "If Outlook is not running please start it and try again",
                 currentMethodName & " - " & ex.ToString,
                 "",
                 CustomDialog.CustomDialogIcons.None,
                 CustomDialog.CustomDialogButtons.OK,
                 CustomDialog.CustomDialogResults.OK)

        Finally
            SetMousePointer(Cursors.Arrow)
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

            Case Is = "ToggleRead"
                If flag Then
                    Me.MenuToggle.Foreground = gForegroundColourEnabled
                Else
                    Me.MenuToggle.Foreground = gForegroundColourDisabled
                End If
                Me.MenuToggle.IsEnabled = flag

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

#Region "Undo variables and logic"

    Private Sub StartAddingToUndoLog()

        ' Prepare for next action log entry
        gUndoLogWasUpdated = True

    End Sub

    Private Sub FinishedAddingToUndoLog()

        If gUndoLogWasUpdated Then
        Else
            Exit Sub
        End If

        gUndoLogWasUpdated = False

        ' Prepare for next action log entry
        gUndoLogIndex += 1
        gUndoLogSubIndex = 0

    End Sub

    Private Sub AddToUndoLog(ByVal Action As String,
                             ByVal FixedIndex As Integer,
                    Optional ByVal EmailEntryID As String = "",
                    Optional ByVal SourceStoreID As String = "",
                    Optional ByVal SourceFolderEntryID As String = "",
                    Optional ByVal TargetStoreID As String = "",
                    Optional ByVal TargetFolderEntryID As String = "",
                    Optional ByVal RecommendationTableRow As ListViewRowClass = Nothing)

        ' Record what happened (so that it can be undone later if necessary)

        If gUndoLogIndex < 0 Then gUndoLogIndex = 0

        EnsureThereIsEnoughtSpaceInUndoLog()

        ' log the action

        Select Case Action

            Case Is = "File"
                gUndoLog(gUndoLogIndex, gUndoLogSubIndex).ActionApplied = ActionType.File
            Case Is = "Delete"
                gUndoLog(gUndoLogIndex, gUndoLogSubIndex).ActionApplied = ActionType.Delete
            Case Is = "ToggleRead"
                gUndoLog(gUndoLogIndex, gUndoLogSubIndex).ActionApplied = ActionType.ToggleRead
            Case Is = "Hide"
                gUndoLog(gUndoLogIndex, gUndoLogSubIndex).ActionApplied = ActionType.Hide

        End Select

        gUndoLog(gUndoLogIndex, gUndoLogSubIndex).FixedIndex = FixedIndex
        gUndoLog(gUndoLogIndex, gUndoLogSubIndex).EmailEntryID = EmailEntryID
        gUndoLog(gUndoLogIndex, gUndoLogSubIndex).SourceStoreID = TargetStoreID
        gUndoLog(gUndoLogIndex, gUndoLogSubIndex).SourceFolderEntryID = SourceFolderEntryID
        gUndoLog(gUndoLogIndex, gUndoLogSubIndex).TargetStoreID = TargetStoreID
        gUndoLog(gUndoLogIndex, gUndoLogSubIndex).TargetFolderEntryID = TargetFolderEntryID
        gUndoLog(gUndoLogIndex, gUndoLogSubIndex).LvrcItem = RecommendationTableRow

        gUndoLogSubIndex += 1

        If Me.MenuUndo.IsEnabled Then
        Else
            MenuOptionEnabled("Undo", True)
        End If

    End Sub

    Private Sub AdjustgEmailTableFromUndo(ByVal newEmailEntryID As String, ByVal lvrci As ListViewRowClass)

        ' redim preserve gEmailTable to add room for one more entry
        ReDim Preserve gEmailTable(gEmailTable.Length)

        Dim newEntry As New StructureOfEmailDetails
        With newEntry
            .sOutlookEntryID = newEmailEntryID
            .sSubject = lvrci.Subject
            .sDateAndTime = lvrci.DateTime
            .sTo = lvrci.xTo
            .sFrom = lvrci.From
            .sOriginalFolderReferenceNumber = lvrci.OriginalFolder
            .sRecommendedFolder1ReferenceNumber = lvrci.RecommendedFolder1
            .sRecommendedFolder2ReferenceNumber = lvrci.RecommendedFolder2
            .sRecommendedFolder3ReferenceNumber = lvrci.RecommendedFolder3
            .sRecommendedFolderFinalReferenceNumber = lvrci.RecommendedFolderFinal
            .sUnRead = lvrci.UnRead
            .sMailBoxName = lvrci.MailBoxName
            .sTrailer = lvrci.Trailer
        End With

        gEmailTable(gEmailTable.Length - 1) = newEntry

    End Sub

    Private Sub AdjustgFinalRecommendationTableFromUndo(ByVal oldEntryID As String, ByVal newEntryID As String, ByVal i1 As Integer, ByVal i2 As Integer)


        Dim isNewEntryAlreadyInThegFinalRecommendationTable As Boolean = False

        Dim i As Integer

        'step one : find the old entry and update it to the new entryID

        For i = 0 To gFinalRecommendationTable.Length - 1

            If gFinalRecommendationTable(i).OutlookEntryID = oldEntryID Then

                gFinalRecommendationTable(i).OutlookEntryID = newEntryID
                Exit For

            End If

        Next

        'step two : check if the new entryID is already in the gFinalRecommendationTable, and if so restore its index

        For i = 0 To gFinalRecommendationTable.Length - 1

            If gFinalRecommendationTable(i).OutlookEntryID = newEntryID Then

                gFinalRecommendationTable(i).Index = gFinalRecommendationTable(i).FixedIndex

                isNewEntryAlreadyInThegFinalRecommendationTable = True

                Exit For

            End If

        Next

        'step three : if the newentryID is not already in the gFinalRecommendationTable then add it in

        If isNewEntryAlreadyInThegFinalRecommendationTable Then

        Else

            ' redim preserve room for one more entry in the gFinalRecommendationTable
            Dim x = gFinalRecommendationTable.Length
            ReDim Preserve gFinalRecommendationTable(x)

            ' restore the final recommendation table entry to the end of the table
            Dim RecomendationTableRow As ListViewRowClass = gUndoLog(i1, i2).LvrcItem.Clone()
            Dim restored As ListViewRowClass = RecomendationTableRow
            restored.OutlookEntryID = newEntryID
            restored.Index = x
            restored.FixedIndex = x

            gFinalRecommendationTable(x) = restored

        End If

        ' the item will be added back to the listview later by in ApplyFilter() 

    End Sub

    Private Sub AdjustgFinalRecommendationTableForHideOrUnHide(ByVal Action As String, ByVal EmailEntryID As String)

        For x = 0 To gFinalRecommendationTable.Length - 1
            If gFinalRecommendationTable(x).OutlookEntryID = EmailEntryID Then
                If Action = "Hide" Then
                    gFinalRecommendationTable(x).Index = -1
                Else
                    gFinalRecommendationTable(x).Index = gFinalRecommendationTable(x).FixedIndex
                End If
                Exit For
            End If
        Next

    End Sub

    Private Sub RestoreFromUndoLog()

        Dim restoredEntryIds As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim undoSelectionSnapshot As SelectionSnapshot = Nothing

        Try

            If gUndoLogIndex < 1 Then
                MenuOptionEnabled("Undo", False)
                gUndoLogIndex = 0
                Exit Try
            End If

            ' Show wait cursor while we do the restore(s)
            SetMousePointer(Cursors.Wait)

            RemoveHandler ListView1.SelectionChanged, AddressOf ListView1_SelectionChanged

            Dim newEmailEntryID As String = ""

            'point back to the last populated log entryId
            gUndoLogIndex -= 1

            Dim i As Integer = 0

            Dim SortOrderResortRequired = False

            While gUndoLog(gUndoLogIndex, i).ActionApplied > 0

                Select Case gUndoLog(gUndoLogIndex, i).ActionApplied

                    Case Is = ActionType.File, ActionType.Delete

                        Dim oldEmailEntryID As String = gUndoLog(gUndoLogIndex, i).EmailEntryID

                        ' move the email back to its original folder (which means swapping source and target details)

                        Dim SourceStoreID As String = gUndoLog(gUndoLogIndex, i).TargetStoreID
                        Dim SourceFolderEntryID As String = gUndoLog(gUndoLogIndex, i).TargetFolderEntryID
                        Dim TargetStoreID As String = gUndoLog(gUndoLogIndex, i).SourceStoreID
                        Dim TargetFolderEntryID As String = gUndoLog(gUndoLogIndex, i).SourceFolderEntryID

                        If (SourceStoreID = TargetStoreID) AndAlso (SourceFolderEntryID = TargetFolderEntryID) Then

                            newEmailEntryID = oldEmailEntryID

                        Else

                            newEmailEntryID = FileMessage(oldEmailEntryID,
                                                          SourceStoreID,
                                                          SourceFolderEntryID,
                                                          TargetStoreID,
                                                          TargetFolderEntryID)

                            AdjustgUndoLog(oldEmailEntryID, newEmailEntryID)

                        End If

                        If Not String.IsNullOrEmpty(newEmailEntryID) Then
                            restoredEntryIds.Add(newEmailEntryID)
                        End If

                        ' we just want to temporarily remove the old entryID from the Main Window until the next refresh - to do this we tweak the gRecommdationTable's index setting it to -1
                        ' we don't however want to add it to the gHiddenEntryIDs list - which would otherwise keep it out of the Main Window for the entire session (or longer)
                        AdjustgFinalRecommendationTableFromUndo(oldEmailEntryID, newEmailEntryID, gUndoLogIndex, i)

                        SortOrderResortRequired = True

                    Case Is = ActionType.ToggleRead

                        RestoreToggleReadStateForEntryID(gUndoLog(gUndoLogIndex, i).EmailEntryID, gUndoLog(gUndoLogIndex, i).LvrcItem)

                        If Not String.IsNullOrEmpty(gUndoLog(gUndoLogIndex, i).EmailEntryID) Then
                            restoredEntryIds.Add(gUndoLog(gUndoLogIndex, i).EmailEntryID)
                        End If

                    Case Is = ActionType.Hide

                        If Not String.IsNullOrEmpty(gUndoLog(gUndoLogIndex, i).EmailEntryID) Then

                            AdjustgFinalRecommendationTableForHideOrUnHide("UnHide", gUndoLog(gUndoLogIndex, i).EmailEntryID)

                            gHiddenEntryIDs.Remove(gUndoLog(gUndoLogIndex, i).EmailEntryID)

                            restoredEntryIds.Add(gUndoLog(gUndoLogIndex, i).EmailEntryID)

                        End If

                        SortOrderResortRequired = True

                End Select

                gUndoLog(gUndoLogIndex, i).ActionApplied = Nothing
                gUndoLog(gUndoLogIndex, i).FixedIndex = Nothing
                gUndoLog(gUndoLogIndex, i).EmailEntryID = Nothing
                gUndoLog(gUndoLogIndex, i).SourceStoreID = Nothing
                gUndoLog(gUndoLogIndex, i).TargetFolderEntryID = Nothing
                gUndoLog(gUndoLogIndex, i).TargetStoreID = Nothing
                gUndoLog(gUndoLogIndex, i).LvrcItem = Nothing

                i += 1

            End While

            If SortOrderResortRequired Then
                ApplyCurrentSortOrderToFinalTable()
            End If

            ApplyFilter() 'force the list view to be rebuilt, adding back in any undone items

            If restoredEntryIds.Count > 0 Then
                undoSelectionSnapshot = New SelectionSnapshot With {
                    .Entries = New List(Of SelectionEntry),
                    .FirstIndex = 0,
                    .HasSelection = False
                }

                For idx As Integer = 0 To ListView1.Items.Count - 1
                    Dim row = TryCast(ListView1.Items(idx), ListViewRowClass)
                    If row Is Nothing Then Continue For

                    If Not String.IsNullOrEmpty(row.OutlookEntryID) AndAlso restoredEntryIds.Contains(row.OutlookEntryID) Then
                        undoSelectionSnapshot.Entries.Add(New SelectionEntry With {
                                                         .OutlookEntryId = row.OutlookEntryID,
                                                         .ChainKey = BuildChainKey(row),
                                                         .Index = row.Index})
                    End If
                Next

                If undoSelectionSnapshot.Entries.Count > 0 Then
                    undoSelectionSnapshot.HasSelection = True
                    undoSelectionSnapshot.FirstIndex = undoSelectionSnapshot.Entries(0).Index
                    For Each entry In undoSelectionSnapshot.Entries
                        If entry.Index < undoSelectionSnapshot.FirstIndex Then
                            undoSelectionSnapshot.FirstIndex = entry.Index
                        End If
                    Next
                End If
            End If

            ListView1.Focus()

            If gUndoLogIndex = 0 Then
                MenuOptionEnabled("Undo", False)
            End If

        Catch ex As Exception

        Finally

            RestoreSelection(undoSelectionSnapshot, SelectionRestoreReason.Undo, undoSelectionSnapshot Is Nothing OrElse Not undoSelectionSnapshot.HasSelection)

            If gAutoChainSelect AndAlso ListView1.SelectedItems.Count > 0 Then
                Dim chainKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each obj In ListView1.SelectedItems
                    Dim row = TryCast(obj, ListViewRowClass)
                    If row IsNot Nothing Then
                        Dim chainKey As String = BuildChainKey(row)
                        If Not String.IsNullOrEmpty(chainKey) Then
                            chainKeys.Add(chainKey)
                        End If
                    End If
                Next

                If chainKeys.Count > 0 Then
                    For i As Integer = 0 To ListView1.Items.Count - 1
                        Dim row = TryCast(ListView1.Items(i), ListViewRowClass)
                        If row Is Nothing Then Continue For
                        Dim chainKey As String = BuildChainKey(row)
                        If Not String.IsNullOrEmpty(chainKey) AndAlso chainKeys.Contains(chainKey) Then
                            If Not ListView1.SelectedItems.Contains(ListView1.Items(i)) Then
                                ListView1.SelectedItems.Add(ListView1.Items(i))
                            End If
                        End If
                    Next
                End If
            End If

            AddHandler ListView1.SelectionChanged, AddressOf ListView1_SelectionChanged
            SetMousePointer(Cursors.Arrow)

        End Try

    End Sub

    Private Sub EnsureThereIsEnoughtSpaceInUndoLog()

        ' Ensure enough space in the undo log; increase size of either or both index by 10% as necessary 

        If gUndoLogIndex > gUndoLogMaxEntries Then
            gUndoLogMaxEntries *= 1.1
            ReDim Preserve gUndoLog(gUndoLogMaxEntries, gUndoLogMaxSubEntries)
        End If

        If gUndoLogSubIndex > gUndoLogMaxSubEntries Then
            gUndoLogMaxSubEntries *= 1.1
            ResizeUndoLogPreservingBothDimensions(gUndoLogMaxEntries, gUndoLogMaxSubEntries) 'redim preserve is only allowed on the first dimensional
        End If

    End Sub

    Private Sub ResizeUndoLogPreservingBothDimensions(ByVal newMaxEntries As Integer, ByVal newMaxSubEntries As Integer)

        If newMaxEntries < 0 Then Throw New ArgumentOutOfRangeException(NameOf(newMaxEntries))
        If newMaxSubEntries < 0 Then Throw New ArgumentOutOfRangeException(NameOf(newMaxSubEntries))

        Dim oldMaxEntries As Integer = gUndoLog.GetUpperBound(0)
        Dim oldMaxSubEntries As Integer = gUndoLog.GetUpperBound(1)

        Dim temp(newMaxEntries, newMaxSubEntries) As StructureOfUndoLog

        Dim entriesToCopy As Integer = Math.Min(oldMaxEntries, newMaxEntries)
        Dim subEntriesToCopy As Integer = Math.Min(oldMaxSubEntries, newMaxSubEntries)

        For i As Integer = 0 To entriesToCopy
            For j As Integer = 0 To subEntriesToCopy
                temp(i, j) = gUndoLog(i, j)
            Next
        Next

        ReDim gUndoLog(newMaxEntries, newMaxSubEntries)

        For i As Integer = 0 To entriesToCopy
            For j As Integer = 0 To subEntriesToCopy
                gUndoLog(i, j) = temp(i, j)
            Next
        Next

    End Sub

#End Region

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

            Case Is = "ToggleRead"
                If Not My.Settings.ConfirmToggle Then
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

            Case "File", "Delete", "Hide", "ToggleRead"

                Dim displayAction As String = strAction.ToLower
                If displayAction = "toggleread" Then
                    displayAction = "toggle the read/unread state Of"
                End If

                Instruction = "Would you like to " & displayAction & " the "

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

                AdditionalDetail = "This prompt can be turned off In the Options Window."

            Case Is = "Open"

                Instruction = "Would you like to open an e-mail?"
                AdditionalDetail = "If you have selected multiple e-mails, only the first one will be opened."

            Case Is = "Undo"

                Instruction = "Would you like to undo your last action?"

            Case Is = "Exit"

                Instruction = "Would you like to exit?"
                AdditionalDetail = "This prompt can be turned off In the Options Window."


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

    Private Sub ActionRequestAgainstAllSelectedItems(ByVal Action As String, ByRef ListView1 As ListView)

        If ListView1.SelectedItems.Count = 0 Then Exit Sub

        Dim selectionSnapshot As SelectionSnapshot = CaptureSelectionSnapshot()

        Me.ForceCursor = True
        Me.Cursor = Cursors.Wait

        gSuppressUpdatesToDetailBox = True

        Try

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

        Select Case Action
            Case Is = "Delete", "File"
                ' in terms of selection restoration, after delete/file/hide we want to restore to the next item after the last one actioned
                RestoreSelection(selectionSnapshot, SelectionRestoreReason.UserDelete, True)

            Case Is = "ToggleRead"
                RestoreSelection(selectionSnapshot, SelectionRestoreReason.ToggleRead, False)
                ListView1.Items.Refresh()

            Case Is = "Hide"
                RestoreSelection(selectionSnapshot, SelectionRestoreReason.Hide, False)
                ListView1.Items.Refresh()
        End Select

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

    Private Sub ToggleReadStateOfASelectedItemIn_TheListView(ByVal index As Integer)

        Dim selectedRow As ListViewRowClass = TryCast(ListView1.Items(index), ListViewRowClass)
        Dim DesiredState As Boolean = (Me.MenuContextToggleRead.Header = "Mark as unread")

        ' Update UI to reflect the new state
        If DesiredState Then
            selectedRow.UnRead = System.Windows.FontWeights.Bold
        Else
            selectedRow.UnRead = System.Windows.FontWeights.Normal
        End If

    End Sub

    Private Function ToggleReadStateOfASelectedItemIn_Outlook(ByVal index As Integer) As Boolean

        ' returns True if a change was made; False otherwise

        Dim result As Boolean = False

        Try

            Dim selectedRow As ListViewRowClass = TryCast(ListView1.Items(index), ListViewRowClass)

            If selectedRow Is Nothing Then
                Exit Try
            End If

            Dim entryId As String = selectedRow.OutlookEntryID
            If String.IsNullOrEmpty(entryId) Then
                Exit Try
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = Nothing

            ' We will try GetItemFromID up to 2 times:
            '  - first with the current session
            '  - on RPC disconnect errors, rebuild session and retry once
            Dim attempt As Integer = 0
            '   Dim errorOccurred As Boolean = False

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

                        ' Let the loop retry GetItemFromID with fresh session
                        mailItem = Nothing
                    Else
                        ' Any other COM error, or second failure: stop processing
                        Exit While
                    End If

                Catch ex As Exception
                    Exit While
                End Try

                attempt += 1

            End While

            If mailItem Is Nothing Then
                ' Could not get mail item, skip this one
                Exit Try
            End If

            ' set the read/unread flag in Outlook
            Try

                ' mailItem.UnRead = Not mailItem.UnRead ' this would be a toggle 
                ' but don't toogle; rather set the state based on the menu text that the user selected

                Dim currentState As Boolean = mailItem.UnRead
                Dim desiredState As Boolean = (Me.MenuContextToggleRead.Header = "Mark as unread")

                If (currentState <> desiredState) Then

                    ' update Outlook
                    Dim action As String = If(desiredState, "ReadGoingToUnread", "UnreadGoingToRead")
                    Call _MainWindow.BlockDuplicateEventProcessing(action, mailItem.EntryID)

                    mailItem.UnRead = desiredState
                    mailItem.Save()

                    result = True

                End If

            Catch comEx As System.Runtime.InteropServices.COMException
                Const RPC_E_SERVER_UNAVAILABLE As Integer = &H800706BA
                Const RPC_E_DISCONNECTED As Integer = &H800706BE

                If comEx.HResult = RPC_E_SERVER_UNAVAILABLE OrElse comEx.HResult = RPC_E_DISCONNECTED Then
                    Exit Try
                End If
            Catch ex As Exception

            Finally

                If mailItem IsNot Nothing Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                    mailItem = Nothing
                End If

            End Try

        Catch ex As Exception
        End Try

        Return result

    End Function
    Private Sub RestoreToggleReadStateForEntryID(ByVal entryId As String, ByVal RecomendationTableRow As ListViewRowClass)

        Try

            gSuppressUpdatesToDetailBox = True

            Try

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

                            ' Let the loop retry GetItemFromID with fresh session
                            mailItem = Nothing

                        End If

                    Catch
                    End Try

                    attempt += 1

                End While

                Try

                    ' set the read/unread flag in Outlook

                    If mailItem IsNot Nothing Then

                        Dim DesiredState As Boolean = (RecomendationTableRow.UnRead = System.Windows.FontWeights.Bold)

                        mailItem.UnRead = DesiredState

                        Dim action As String = If(mailItem.UnRead, "ReadGoingToUnread", "UnreadGoingToRead")
                        Call _MainWindow.BlockDuplicateEventProcessing(action, mailItem.EntryID)

                        mailItem.Save()

                        ' Update UI to the state found in the undo log

                        For Each item In ListView1.Items

                            Dim updatedRow As ListViewRowClass = CType(item, ListViewRowClass)

                            If updatedRow.OutlookEntryID = entryId Then

                                If DesiredState Then
                                    updatedRow.UnRead = System.Windows.FontWeights.Bold
                                Else
                                    updatedRow.UnRead = System.Windows.FontWeights.Normal
                                End If

                                Exit For

                            End If

                        Next

                    Else
#If DEBUG Then
                        ' in general this should not happen,
                        ' However, if the user used Outlook to delete or move an email then 
                        ' its original EntryID, recorded in the undoLog, will no longer be valid
                        ' we will ignore this undo request to toggle its read/unread status
                        Console.WriteLine("Failed to retrieve e-mail from Outlook, EntryId = " & entryId)
#End If


                    End If

                Catch comEx As System.Runtime.InteropServices.COMException
                    Const RPC_E_SERVER_UNAVAILABLE As Integer = &H800706BA
                    Const RPC_E_DISCONNECTED As Integer = &H800706BE

                    If comEx.HResult = RPC_E_SERVER_UNAVAILABLE OrElse comEx.HResult = RPC_E_DISCONNECTED Then
                        Exit Try
                    End If

                Catch ex As Exception

                Finally

                    If mailItem IsNot Nothing Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                        mailItem = Nothing
                    End If

                End Try

            Finally

                gSuppressUpdatesToDetailBox = False

            End Try

        Catch ex As Exception

        End Try

    End Sub

    Private Sub AdjustgEmailTableForFileOrDelete(ByVal oldEmailEntryID As String, ByVal newEmailEntryID As String, ByVal newFolderEntryID As String)

        Dim newFolderReferenceNumber As Int32 = -1

        For i As Int32 = 0 To gFolderTable.Length - 1
            If gFolderTable(i).EntryID = newFolderEntryID Then
                newFolderReferenceNumber = i
                Exit For
            End If
        Next

        ' update the gEmailTable's with their newEmailEntryID and new folder reference number
        For i As Integer = 0 To gEmailTable.Length - 1
            If gEmailTable(i).sOutlookEntryID = oldEmailEntryID Then
                gEmailTable(i).sOutlookEntryID = newEmailEntryID
                gEmailTable(i).sOriginalFolderReferenceNumber = newFolderReferenceNumber
                Exit For
            End If
        Next

    End Sub


    Private Sub AdjustgUndoLog(ByVal oldEntryID As String, ByVal newEntryID As String)

        ' when we delete or file an email it gets another EntryID and moves to another folder

        ' update the undoLog to reference a deleted or filed email by its new EntryID in the future

        For x As Integer = 0 To gUndoLogIndex - 1

            For y As Integer = 0 To gUndoLogMaxSubEntries - 1

                If gUndoLog(x, y).EmailEntryID IsNot Nothing Then

                    If gUndoLog(x, y).EmailEntryID = oldEntryID Then
                        gUndoLog(x, y).EmailEntryID = newEntryID
                        gUndoLog(x, y).LvrcItem.OutlookEntryID = newEntryID
                        Exit For
                    End If

                Else
                    Exit For
                End If
            Next

        Next

    End Sub

    Private Sub AdjustgFinalRecommendationTableFromFileOrDelete(ByVal oldEntryID As String, ByVal newEntryID As String)

        Dim i As Integer

        For i = 0 To gFinalRecommendationTable.Length - 1

            If gFinalRecommendationTable(i).OutlookEntryID = oldEntryID Then

                gFinalRecommendationTable(i).OutlookEntryID = newEntryID

                Exit For

            End If

        Next

    End Sub
    Private Sub ActionRequest_Worker(ByVal Action As String, ByRef SelectedEntries() As Integer, ByVal NumberOfSelectedItems As Integer, ByRef ListView1 As ListView)

        Try

            'action requests

            Dim IndexToAction As Integer

            StartAddingToUndoLog()

            Dim ListOfOldEntryIDsThatHaveBeenRemoved As New List(Of String)

            Dim ActionsTaken As ActionType = ActionType.None

            For z As Integer = 0 To NumberOfSelectedItems - 1

                IndexToAction = SelectedEntries(z)

                Dim ClonedListViewItem As ListViewRowClass = ListView1.Items(IndexToAction).Clone()

                Select Case Action

                    Case "File", "Delete"

                        ActionsTaken = ActionType.File

                        Dim oldEmailEntryID As String = ClonedListViewItem.OutlookEntryID

                        Dim oldgFolderTableIndex As Integer = ClonedListViewItem.OriginalFolder

                        Dim SourceStoreID As String = gFolderTable(oldgFolderTableIndex).StoreID
                        Dim SourceFolderEntryID As String = gFolderTable(oldgFolderTableIndex).EntryID

                        Dim newgFolderTableIndex As Integer

                        If Action = "File" Then
                            newgFolderTableIndex = gPickFromContextMenuOverride
                        Else
                            ' Action = "Delete"
                            ' Delete: choose a per‑store Deleted/Trash folder
                            Dim sourceStoreIdForDelete As String = SourceStoreID
                            newgFolderTableIndex = GetDeleteFolderIndexForStore(sourceStoreIdForDelete)
                        End If

                        Dim TargetStoreID As String = gFolderTable(newgFolderTableIndex).StoreID
                        Dim TargetFolderEntryID As String = gFolderTable(newgFolderTableIndex).EntryID

                        Dim newEmailEntryID As String

                        If (SourceStoreID = TargetStoreID) AndAlso (SourceFolderEntryID = TargetFolderEntryID) Then

                            ' move not needed
                            newEmailEntryID = oldEmailEntryID

                        Else

                            newEmailEntryID = FileMessage(oldEmailEntryID,
                                                          SourceStoreID,
                                                          SourceFolderEntryID,
                                                          TargetStoreID,
                                                          TargetFolderEntryID)

                            AdjustgUndoLog(oldEmailEntryID, newEmailEntryID)
                            AdjustgFinalRecommendationTableFromFileOrDelete(oldEmailEntryID, newEmailEntryID)
                            AdjustgEmailTableForFileOrDelete(oldEmailEntryID, newEmailEntryID, TargetFolderEntryID)
                            ListOfOldEntryIDsThatHaveBeenRemoved.Add(oldEmailEntryID)

                        End If

                        AddToUndoLog(Action,
                                     ClonedListViewItem.FixedIndex,
                                     newEmailEntryID,
                                     SourceStoreID,
                                     SourceFolderEntryID,
                                     TargetStoreID,
                                     TargetFolderEntryID,
                                     gFinalRecommendationTable(ClonedListViewItem.Index))

                        AdjustgFinalRecommendationTableForHideOrUnHide("Hide", oldEmailEntryID)

                        If newEmailEntryID <> oldEmailEntryID Then
                            AdjustgFinalRecommendationTableForHideOrUnHide("Hide", newEmailEntryID)
                        End If

                    Case "ToggleRead"

                        If ToggleReadStateOfASelectedItemIn_Outlook(IndexToAction) Then

                            ActionsTaken = ActionType.ToggleRead

                            AddToUndoLog(Action, ClonedListViewItem.FixedIndex, ClonedListViewItem.OutlookEntryID,,,,, ClonedListViewItem)
                            ToggleReadStateOfASelectedItemIn_TheListView(IndexToAction)

                        End If

                    Case "Hide"

                        ActionsTaken = ActionType.Hide

                        AddToUndoLog(Action, ClonedListViewItem.FixedIndex, ClonedListViewItem.OutlookEntryID,,,,, ClonedListViewItem)
                        AdjustgFinalRecommendationTableForHideOrUnHide("Hide", ClonedListViewItem.OutlookEntryID)

                        gHiddenEntryIDs.Add(ClonedListViewItem.OutlookEntryID)

                End Select

            Next z

            FinishedAddingToUndoLog()

            For Each entryId In ListOfOldEntryIDsThatHaveBeenRemoved
                LiftSuspensionOnRemovedEntryId(entryId)
            Next

            If ActionsTaken = ActionType.File OrElse ActionsTaken = ActionType.Hide Then '(this covers a delete as well)
                ApplyFilter()
            End If

        Catch ex As Exception

            Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

            Call ShowMessageBox("FileFriendly",
                     CustomDialog.CustomDialogIcons.Stop,
                     "Unexpected Error!",
                     "FileFriendly has encountered an unexpected Error.",
                     currentMethodName & " - " & ex.ToString,
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

            Dim outlookWasRunning As Boolean = IsOutlookProcessRunning()

            If Not repairOnly AndAlso Not outlookWasRunning Then

                Dim header As String = "FileFriendly - Start Outlook?"
                Dim instruction As String = vbCrLf &
                    "Microsoft Outlook is not running." & vbCrLf & vbCrLf &
                    "FileFriendly needs Outlook to be running to help file your e-mails." & vbCrLf & vbCrLf &
                    "You can either start Outlook manually or have FileFriendly start it for you now." & vbCrLf & vbCrLf & "Would you like FileFriendly to start  Outlook for you now?"
                Dim detail As String =
                    "If you choose 'Yes', FileFriendly will automatically start Outlook." & vbCrLf & vbCrLf &
                    "If you choose 'No', FileFriendly will close unless you have manually started Outlook yourself."

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

                If response = CustomDialog.CustomDialogResults.Yes Then

                    If IsOutlookProcessRunning() Then
                        ' Outlook was started in the meantime 
                        outlookWasRunning = True
                        GoTo OutlookIsNowRunning
                    End If

                Else

                    If IsOutlookProcessRunning() Then
                        ' Outlook was started in the meantime 
                        outlookWasRunning = True
                        GoTo OutlookIsNowRunning
                    End If

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

            End If

            If outlookWasRunning Then GoTo OutlookIsNowRunning

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

            outlookWasRunning = startedOk

OutlookIsNowRunning:

            If Not outlookWasRunning Then Return False

            ScheduleRefreshGrid()

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

            Dim currentMethodName As String = System.Reflection.MethodBase.GetCurrentMethod().Name

            Call ShowMessageBox("FileFriendly - Outlook error",
                 CustomDialog.CustomDialogIcons.Stop,
                 "Unexpected Error!",
                 "FileFriendly has encountered an unexpected error.",
                 currentMethodName & " - " & ex.ToString,
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

        If (LastState = Windows.WindowState.Minimized) OrElse (LastState = Windows.WindowState.Maximized) Then
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

    Private Sub ApplyCurrentSortOrderToFinalTable()

        If gFinalRecommendationTable Is Nothing OrElse gFinalRecommendationTable.Length = 0 Then Return

        Dim column As String
        Dim direction As FinalRecommendationTableSorter.MySortOrder

        column = gCurrentSortOrder
        direction = If(gCurrentSortDirection = ListSortDirection.Descending,
                           FinalRecommendationTableSorter.MySortOrder.Descending,
                           FinalRecommendationTableSorter.MySortOrder.Ascending)

        Dim sorter As New FinalRecommendationTableSorter With {
          .PrimaryColumnToSort = column,
          .SortOrder = direction
        }

        Array.Sort(gFinalRecommendationTable, sorter)

        UpdateSortHeaderGlyph()

    End Sub

    Private Sub UpdateSortHeaderGlyph()

        If Not Dispatcher.CheckAccess() Then
            Dispatcher.BeginInvoke(New Action(AddressOf UpdateSortHeaderGlyph))
            Return
        End If

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

                    If String.Equals(header, gCurrentSortOrder, StringComparison.OrdinalIgnoreCase) Then
                        If gCurrentSortDirection = ListSortDirection.Ascending Then
                            direction = ListSortDirection.Descending
                        Else
                            direction = ListSortDirection.Ascending
                        End If
                    ElseIf headerClicked IsNot _lastheaderClicked Then
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

                    gCurrentSortDirection = direction
                    gCurrentSortOrder = header
                    SetListViewItem(gFinalRecommendationTable)

                    If direction = ListSortDirection.Ascending Then
                        headerClicked.Column.HeaderTemplate = TryCast(Resources("HeaderTemplateArrowUp"), DataTemplate)
                    Else
                        headerClicked.Column.HeaderTemplate = TryCast(Resources("HeaderTemplateArrowDown"), DataTemplate)
                    End If

                    ' Remove arrow _from previously sorted header
                    If _lastheaderClicked IsNot Nothing AndAlso _lastheaderClicked IsNot headerClicked Then
                        _lastheaderClicked.Column.HeaderTemplate = Nothing
                    End If

                    _lastheaderClicked = headerClicked
                    _lastDirection = direction

                    UpdateSortHeaderGlyph()

                End If

            End If

            ApplyFilter()

            ListView1.Items.Refresh()

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

    ' Thread-safe helpers for gSuppressEventForEntryIds / _suspensionTimers.

    Private Sub WaitUntilThereAreNoLongerAnySuprressedEvents()

        While gSuppressEventForEntryIds.Count() > 0
            Thread.Sleep(100)
            DoEvents()
        End While

    End Sub

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

#Region "Event Loop Prevention"

#End Region

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
                        Dim exUser As Microsoft.Office.Interop.Outlook.ExchangeUser = Nothing
                        Try
                            If mailItem.Sender IsNot Nothing Then
                                If String.Equals(mailItem.SenderEmailType, "SMTP", StringComparison.OrdinalIgnoreCase) Then
                                    friendlyFrom = mailItem.SenderEmailAddress
                                Else
                                    exUser = TryCast(mailItem.Sender.GetExchangeUser(), Microsoft.Office.Interop.Outlook.ExchangeUser)
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
                        Finally
                            If exUser IsNot Nothing Then
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(exUser)
                                exUser = Nothing
                            End If
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
            ' 1. the email is deleted or removed _from a monitored folder via FileFriendly
            ' 2. the email is deletes or removed _from a monitored folder via Outlook

            ' However, in neither scenario do we know the EmailEntryID of the deleted / removed email
            ' I had originally tried to track the EntryIDs in the listview by taking a snapshot of them before and after the removal and then comparing the two snapshots to deduce which EmailEntryID(s) had been removed
            ' However, this approach does not work in second case (removal in Outlook) as the listview doesn't change just because an email is deleted in Outlook

            ' Accordingly, on a removal the program will wait for the last OnItemRemove event from Outlook and then refresh the Main Window

            ' wait until the last OnItemRemove event from Outlook 

            If _mainWindow.GetSuppressedEventCount() > 1 Then Return

            _mainWindow.WaitUntilThereAreNoLongerAnySuprressedEvents()

            ' refresh the Main Window 

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

                Dim entryId As String = ""

                Try
                    entryId = mailItem.EntryID
                Catch
                    Return
                End Try

                Dim action As String = If(mailItem.UnRead, "ReadGoingToUnread", "UnreadGoingToRead")
                If _mainWindow.BlockDuplicateEventProcessing(action, entryId) Then Return

                _mainWindow.ScheduleRefreshGrid()

            Catch ex As Exception

            Finally
                Try
                    If mailItem IsNot Nothing Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem)
                        mailItem = Nothing
                    End If
                Catch
                End Try

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

    Implements System.Collections.Generic.IComparer(Of MainWindow.StructureOfEmailDetails)

    Public Enum SortOrder As Integer
        None = 0
        Ascending = 1
        Descending = 2
    End Enum

    Private Shared ReadOnly SubjectComparer As StringComparer = StringComparer.Ordinal
    Private Shared ReadOnly DateComparer As System.Collections.Generic.IComparer(Of DateTime) = System.Collections.Generic.Comparer(Of DateTime).Default
    Public Shared ReadOnly SubjectThenDateAsc As New EMailTableSorter(SortOrder.Ascending, SortOrder.Descending)

    Private ReadOnly _primaryOrder As SortOrder
    Private ReadOnly _secondaryOrder As SortOrder

    Private Sub New(primaryOrder As SortOrder, secondaryOrder As SortOrder)
        _primaryOrder = primaryOrder
        _secondaryOrder = secondaryOrder
    End Sub

    Public Function Compare(ByVal x As MainWindow.StructureOfEmailDetails, ByVal y As MainWindow.StructureOfEmailDetails) As Integer Implements System.Collections.Generic.IComparer(Of MainWindow.StructureOfEmailDetails).Compare
        Dim result As Integer

        result = SubjectComparer.Compare(x.sSubject, y.sSubject)
        If result <> 0 Then
            Return If(_primaryOrder = SortOrder.Ascending, result, -result)
        End If

        result = SubjectComparer.Compare(x.sTrailer, y.sTrailer)
        If result <> 0 Then
            Return If(_primaryOrder = SortOrder.Ascending, result, -result)
        End If

        result = DateComparer.Compare(x.sDateAndTime, y.sDateAndTime)
        Return If(_secondaryOrder = SortOrder.Ascending, result, -result)
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

    Public Sub New()

        ObjectCompare = New Comparer(System.Globalization.CultureInfo.CurrentCulture)

    End Sub

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
                    ' Same _subject: compare by DateTime, most recent first
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