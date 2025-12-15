Imports System.Diagnostics.Eventing.Reader
Imports System.Management
Imports System.Security.Policy
Imports System.Threading
Imports System.Windows


Class MainWindow
    Inherits Window

    ' Include reference to Outlook 12 Object

    Private gForegroundColourAlert As System.Windows.Media.SolidColorBrush
    Private gForegroundColourEnabled As System.Windows.Media.SolidColorBrush
    Private gForegroundColourDisabled As System.Windows.Media.SolidColorBrush

    Private gThresholdForReportingProgressOnTheProgressBar As Double

    Private gSupressUpdatesToDetailBox As Boolean = False

    Private gEmailExcludeList As String = ""

    Private FileCommand As System.Windows.Forms.TextBox = New System.Windows.Forms.TextBox

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

    Private EmailTableIndex As Integer = 0
    Private EmailTable(1) As StructureOfEmailDetails

    Private lTotalEMails As Integer = 0
    Private lTotalEMailsToBeReviewed As Integer = 0
    Private lTotalRecommendations As Integer = 0

    Structure StructureOfFolderOccurances
        Dim FolderReference As Int16
        Dim FolderReferenceCount As Int16
    End Structure
    Structure StructureOfUniqueSubjectsDetails
        Dim SubjectAndTrailer As String
        Dim TotalOccurances As Int16
        Dim Occurances() As StructureOfFolderOccurances
    End Structure

    Private UniqueSubjectsTableIndex As Integer = 0
    Private UniqueSubjectsTable(1) As StructureOfUniqueSubjectsDetails

    Private gOriginalWidthForm, gOriginalWidthSubject, gOriginalWidthTo, gOriginalWidthFrom, gOriginalWidthDate As Integer

    Private gViewSent As Boolean = True
    Private gViewInbox As Boolean = True
    Private gViewAll As Boolean = True
    Private gViewRead As Boolean = True
    Private gViewUnRead As Boolean = True

    Private gFinalRecommendationTable(1) As ListViewRowClass

    Private gIsRefreshing As Boolean = False
    Private gCancelRefresh As Boolean = False

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
    Private _mailboxCount As Integer = 0

    ' Distinct mailbox names encountered during this refresh
    Private ReadOnly _mailboxNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Private gOriginalTabControl2Height As Integer

    Private oApp As Microsoft.Office.Interop.Outlook.Application
    Private oNS As Microsoft.Office.Interop.Outlook.NameSpace
    Private oMailItem As Microsoft.Office.Interop.Outlook.MailItem
    Private oTargetFolder As Microsoft.Office.Interop.Outlook.MAPIFolder

    ' MD5 instance used by ComputeTrailerHash
    Private md5Obj As New System.Security.Cryptography.MD5CryptoServiceProvider

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

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

        Try

            'Ensure Settings are kept thru an upgrade

            '**********************************************
            '************** DeleteHiddenMessage()  '*** only uncomment for testing security
            '***********************************************

            ' Verify Outlook is available using Outlook object model instead of Redemption
            Try
                Dim testNs As Microsoft.Office.Interop.Outlook.NameSpace = oApp.Session
                Dim version As String = testNs.Application.Version
                If String.IsNullOrEmpty(version) Then
                    Throw New Exception("Outlook version is empty")
                End If
            Catch
                MsgBox("FileFriendly has encountered a problem and cannot continue." & vbCrLf & vbCrLf &
                       "It appears that Microsoft Outlook is not installed or accessible on this computer." & vbCrLf & vbCrLf &
                       "FileFriendly requires Outlook to be able to run.",
                       MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "FileFriendly - Critical Error")
                End
            End Try


            gForegroundColourAlert = System.Windows.Media.Brushes.Red
            gForegroundColourEnabled = Me.MenuAbout.Foreground
            gForegroundColourDisabled = System.Windows.Media.Brushes.Gray

            Dim a As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim appVersion As Version = a.GetName().Version
            Dim appVersionString As String = appVersion.ToString

            ReDim gDockSound(My.Resources.dock.Length)

            My.Resources.dock.Read(gDockSound, 0, My.Resources.dock.Length)

            If My.Settings.ApplicationVersion <> appVersion.ToString Then
                My.Settings.Upgrade()
                My.Settings.ApplicationVersion = appVersionString
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
            gRefreshAll = My.Settings.ScanAll

            'gViewRead = My.Settings.ViewRead
            'gViewUnRead = My.Settings.ViewUnRead
            'gViewAll = My.Settings.ViewAll
            'gViewInbox = My.Settings.ViewInbox
            'gViewSent = My.Settings.ViewSent

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

            RefreshGrid()

            ReDim ActionLog(ActionLogMaxSubEntries, ActionLogMaxSubEntries)

            My.Settings.Save()

            MemoryManagement.FlushMemory()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.OkOnly, "FileFriendly - Loading Error")
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

                ShowMessageBox("FileFriendly - Note",
               CustomDialog.CustomDialogIcons.Information,
               "FileFriendly is already running",
               "Only one instance of FileFriendly can be run at once.",
               "The orginal instance of Filefriendly will remain running, but a new one will not be started",
               "",
               CustomDialog.CustomDialogIcons.None,
               CustomDialog.CustomDialogButtons.OK,
               CustomDialog.CustomDialogResults.OK)

                End

            End If

        Catch ex As Exception
        End Try

    End Sub

    Private Sub MainWindow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing

        SetProcessPriorities("Shutdown")

        On Error Resume Next

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
            MsgBox(ex.ToString)
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

                    Try

                        oApp = CreateObject("Outlook.Application")
                        oNS = oApp.GetNamespace("MAPI")

                        Try ' added in v1.3 to allow access to exchange servers
                            oNS.Logon(oNS.CurrentProfileName.ToString, "", False, False)
                        Catch ex1 As Exception
                            Try
                                oNS.Logon("Outlook", "", False, False)
                            Catch ex2 As Exception
                                Try
                                    oNS.Logon("Outlook", "", False, True)
                                Catch ex3 As Exception
                                End Try
                            End Try

                        End Try

                    Catch ex As Exception

                        ShowMessageBox("FileFriendly",
                        CustomDialog.CustomDialogIcons.Information, "FileFriendly cannot access Microsoft Outlook on this machine.",
                        vbCrLf & "If Outlook is not installed, please be aware FileFriendly needs Outlook to run." & vbCrLf &
                        vbCrLf & "If Outlook is installed running Microsoft's free scanpst program may resolve the problem.",
                        ex.ToString,
                        gGithubWebPage, CustomDialog.CustomDialogIcons.None,
                        CustomDialog.CustomDialogButtons.OK)

                        Me.Close()
                        End

                    End Try

                    EnsureOutlookIsRunning()

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
                    'myProcess.PriorityClass = ProcessPriorityClass.BelowNormal
                    'OutlookProcess.PriorityClass = ProcessPriorityClass.BelowNormal

                    myProcess.PriorityClass = ProcessPriorityClass.Normal
                    OutlookProcess.PriorityClass = ProcessPriorityClass.Normal

                Case Is = "End Outlook Review"
                    OutlookProcess.PriorityClass = ProcessPriorityClass.Normal

                Case Is = "End Review"
                    myProcess.PriorityClass = ProcessPriorityClass.Normal

                Case Is = "Shutdown"
                    OutlookProcess.PriorityClass = ProcessPriorityClass.Normal
                    myProcess.Dispose()
                    OutlookProcess.Dispose()

            End Select

        Catch ex As Exception

        End Try

    End Sub

#Region "List View Stuff"

    Delegate Sub SetCursorCallback(ByVal [CursorType] As System.Windows.Input.Cursor)
    Private Sub SetCursor(ByVal [CursorType] As System.Windows.Input.Cursor)
        Me.Cursor = [CursorType]
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
        gViewAll = gRefreshAll

        MenuViewInbox.IsChecked = gRefreshInbox
        MenuViewSent.IsChecked = gRefreshSent
        MenuViewAll.IsChecked = gRefreshAll

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

        ' Reset mailbox tracking for this refresh
        _mailboxNames.Clear()
        _mailboxCount = 0

    End Sub

    Delegate Sub FinalizeLoadCallback()
    Private Sub FinalizeLoad()

        ApplyFilter()

        ' Adjust Mailbox column after items are loaded
        UpdateMailboxColumnVisibility()
        RecalculateListViewColumnWidths()

        gIsRefreshing = False
        gCancelRefresh = False
        UpdateRefreshMenuState()

        'MenuOptionEnabled("Refresh", True)
        MenuOptionEnabled("Options", True)

        If ListView1.Items.Count > 0 Then

            Me.MenuActions.Foreground = gForegroundColourEnabled

            MenuOptionEnabled("Open", True)
            MenuOptionEnabled("Hide", True)
            MenuOptionEnabled("Delete", True)
            If ActionLogIndex > 0 Then MenuOptionEnabled("Undo", True)
            MenuOptionEnabled("View", True)

            'Force the first entry to be selected
            Me.ListView1.Focus()
            ListView1.SelectedIndex = 0

            If gAutoChainSelect Then SelectAllMembersOfAnEmailChain()

        Else

            If gRefreshAll Or gRefreshSent Or gRefreshInbox Then
                Me.MenuRefresh.Foreground = gForegroundColourEnabled
                Me.MenuActions.Foreground = gForegroundColourEnabled
            Else
                Me.MenuRefresh.Foreground = gForegroundColourAlert
                Me.MenuActions.Foreground = gForegroundColourAlert
            End If

        End If

        If My.Settings.SoundScanComplete Then
            Beep()
        End If

    End Sub


    Dim Timer1 As System.Windows.Threading.DispatcherTimer
    Delegate Sub TimerControlCallback(ByVal [Command] As String, ByVal [TimeTakenToProcessFiles] As Integer)

    ' Add this field to hold the Tick handler delegate
    Private ReadOnly Timer1TickHandler As New EventHandler(AddressOf Timer1_Tick)

    Private Sub TimerControl(ByVal [Command] As String, ByVal [TimeTakenToProcessFiles] As Integer)

        If [Command] = "Start" Then

            Dim UpdateProgressBarAfterThisManyMilliSeconds As Double = 300

            Dim RemainingRecords As Double = ProgressBar1.Maximum - ProgressBar1.Value

            Dim EstimatedTimeToComplete As Double = [TimeTakenToProcessFiles] * My.Settings.RatioOfRecommendationToProcessingTime

            ' calculate how far the progress bar should progress in a second
            lProgressBarUpdateUsedByTimer = RemainingRecords / (EstimatedTimeToComplete / 1000)
            ' calculate how far the progress bar should progress in each fraction of a second
            lProgressBarUpdateUsedByTimer *= (UpdateProgressBarAfterThisManyMilliSeconds / 1000)


            Timer1 = New System.Windows.Threading.DispatcherTimer(Windows.Threading.DispatcherPriority.Send)
            ' Use the cached delegate instead of AddressOf directly
            AddHandler Timer1.Tick, Timer1TickHandler
            Timer1.IsEnabled = True
            Timer1.Interval = TimeSpan.FromMilliseconds(250)
            Timer1.Start()

        Else

            Timer1.Stop()
            Timer1.IsEnabled = False
            ' Use the cached delegate here as well
            RemoveHandler Timer1.Tick, Timer1TickHandler
            Timer1 = Nothing

        End If

    End Sub

    Private lProgressBareRefreshingThresholdCounter As Integer = 0
    Private lProgressBarUpdateUsedByTimer As Double
    Friend Sub Timer1_Tick()

        If (ProgressBar1.Value + lProgressBarUpdateUsedByTimer) <= ProgressBar1.Maximum Then
            ProgressBar1.Value += lProgressBarUpdateUsedByTimer
        Else
            ProgressBar1.Value = ProgressBar1.Maximum
        End If

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
            ' 3 orginal email of chain
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
            MenuOptionEnabled("Refresh", (gRefreshInbox Or gRefreshSent Or gRefreshAll))
        End If
    End Sub


    Delegate Sub SetListViewItemCallback(ByVal FinalRecommendationTable As ListViewRowClass())

    Private Sub SetListViewItem(ByVal FinalRecommendationTable As ListViewRowClass())

        Try

            If FinalRecommendationTable.Length = 0 Then Exit Sub

            Dim lFirstEntry As Boolean = True
            Dim lCurrentSubjectPlusTrailer As String = ""
            Dim lPreviousSubjectPlusTrailer As String = ""

            Dim LineCount As Integer = 0

            ListView1.Items.Clear()

            ResetEMailChainRelationShips(FinalRecommendationTable)

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

            MsgBox(ex.TargetSite.Name & " - " & ex.ToString)

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

        'special case to deal with final entry
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

            If Not (gRefreshInbox Or gRefreshSent Or gRefreshAll) Then
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
                ListView1.SelectedIndex = 0
            End If

            UpdateDetails()

        Catch ex As Exception
            ' Optional: log ex.ToString()
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
        'MenuOptionEnabled("View", False)

    End Sub
    Private Sub RefreshGrid()

        Try

            Me.ListView1.Visibility = Windows.Visibility.Hidden

            ' Remove arrow from previously sorted header
            If _lastheaderClicked IsNot Nothing Then
                _lastheaderClicked.Column.HeaderTemplate = Nothing
            End If

            BlankOutDetails()

            'start refresh process
            Dim t As System.Threading.Thread
            t = New System.Threading.Thread(AddressOf Me.RefreshBackGroundTask)
            t.IsBackground = True
            t.Start()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Collection_of_folders_to_exclude As System.Collections.Specialized.StringCollection = New System.Collections.Specialized.StringCollection
    Private Collection_of_folders_to_exclude_is_empty As Boolean = True

    Private Sub RefreshBackGroundTask()

        'Dim swOverall As Stopwatch = New Stopwatch : swOverall.Start()

        Try

            MemoryManagement.FlushMemory()

            Me.Dispatcher.Invoke(New BeginLoadCallback(AddressOf BeginLoad), New Object() {})

            ' if cancelled before we even start, honour it
            If gCancelRefresh Then GoTo CleanExit

            Me.Dispatcher.Invoke(New SetCursorCallback(AddressOf SetCursor), New Object() {Cursors.Wait})

            SetProcessPriorities("Start Outlook Review")

            Me.Dispatcher.Invoke(New SetToolTipCallback(AddressOf SetToolTip), New Object() {"Folders are being reviewed"})

            Collection_of_folders_to_exclude = My.Settings.ExcludedScanFolders 'list of all folders to be excluded from scan

            Collection_of_folders_to_exclude_is_empty = (Collection_of_folders_to_exclude Is Nothing)

            Me.Dispatcher.Invoke(New ShowFoldersCallback(AddressOf ShowFolders), New Object() {})

            gMinimizeMaximizeAllowed = True

            'Dim swxx As Stopwatch = New Stopwatch : swxx.Start()
            FindAllFolders()
            'swxx.Stop() : Console.WriteLine(swxx.ElapsedMilliseconds)

            Me.Dispatcher.Invoke(
            New SetFolderNameTextCallback(AddressOf SetFoldersNameText),
            New Object() {String.Format("DEBUG: Inbox={0} Sent={1} All={2} ToReview={3}",
                                gRefreshInbox, gRefreshSent, gRefreshAll, lTotalEMailsToBeReviewed)})


            Me.Dispatcher.Invoke(New ShowFoldersCallback(AddressOf ShowFolders), New Object() {})

            If gRefreshInbox OrElse gRefreshSent OrElse gRefreshAll Then

                If lTotalEMailsToBeReviewed > 0 Then

                    Me.Dispatcher.Invoke(New SetFolderNameTextCallback(AddressOf SetFoldersNameText), New Object() {"Reviewing " & lTotalEMailsToBeReviewed.ToString("#,#", System.Globalization.CultureInfo.InvariantCulture) & " of " & lTotalEMails.ToString("#,#", System.Globalization.CultureInfo.InvariantCulture) & " e-mails"})

                    Me.Dispatcher.Invoke(New SetProgressBarVisableCallback(AddressOf SetProgressBarVisable), New Object() {Windows.Visibility.Visible})

                    Dim sw As Stopwatch = New Stopwatch : sw.Start()

                    ProcessAllFolders()

                    If Not gCancelRefresh Then

                        Collection_of_folders_to_exclude = Nothing
                        Collection_of_folders_to_exclude_is_empty = True

                        SetProcessPriorities("End Outlook Review")

                        MemoryManagement.FlushMemory()

                        sw.Stop()
                        Dim TimeTakenToProcessFiles As Integer = sw.ElapsedMilliseconds
                        sw.Reset()
                        sw.Start()

                        If TimeTakenToProcessFiles > 2000 Then Me.Dispatcher.Invoke(New TimerControlCallback(AddressOf TimerControl), New Object() {"Start", TimeTakenToProcessFiles})

                        EstablishRecomendations()

                        UpdateListView()

                        sw.Stop()

                        If TimeTakenToProcessFiles > 2000 Then Me.Dispatcher.Invoke(New TimerControlCallback(AddressOf TimerControl), New Object() {"Stop", 0})

                        Me.Dispatcher.Invoke(New SetProgressBarVisableCallback(AddressOf SetProgressBarVisable), New Object() {Windows.Visibility.Hidden})

                        Me.Dispatcher.Invoke(New SetFolderNameTextCallback(AddressOf SetFoldersNameText), New Object() {Format(lTotalRecommendations, "###,####,###") & " e-mails"})

                        Me.Dispatcher.Invoke(New SetProgressBarVisableCallback(AddressOf SetProgressBarVisable), New Object() {Windows.Visibility.Hidden})

                        Me.Dispatcher.Invoke(New SetFolderNameTextCallback(AddressOf SetFoldersNameText), New Object() {Format(lTotalRecommendations, "###,####,###") & " e-mails"})

                        Dim TimeTakenToMakeRecommendations As Integer = sw.ElapsedMilliseconds

                        My.Settings.RatioOfRecommendationToProcessingTime = TimeTakenToMakeRecommendations / TimeTakenToProcessFiles

                        My.Settings.Save()

                        sw = Nothing

                    Else

                        ' Cancelled after ProcessAllFolders; show partial results if any
                        Collection_of_folders_to_exclude = Nothing
                        Collection_of_folders_to_exclude_is_empty = True
                        SetProcessPriorities("End Outlook Review")
                        MemoryManagement.FlushMemory()

                        ' hide and reset progress bar on cancel
                        Me.Dispatcher.Invoke(New SetProgressBarVisableCallback(AddressOf SetProgressBarVisable),
                                             New Object() {Windows.Visibility.Hidden})
                        Me.Dispatcher.Invoke(New SetProgressBarValueCallback(AddressOf SetProgressBarValue),
                                             New Object() {0.0R})

                        ' reset status-line email count on cancel
                        Me.Dispatcher.Invoke(New SetFolderNameTextCallback(AddressOf SetFoldersNameText),
                                             New Object() {"0 e-mails"})

                    End If

                End If

                Me.Dispatcher.Invoke(New SetToolTipCallback(AddressOf SetToolTip), New Object() {"Done"})

                Me.Dispatcher.Invoke(New SetCursorCallback(AddressOf SetCursor), New Object() {Cursors.Arrow})

                SetProcessPriorities("End Review")

            Else

                Collection_of_folders_to_exclude = Nothing
                Collection_of_folders_to_exclude_is_empty = True

                SetProcessPriorities("End Outlook Review")

                MemoryManagement.FlushMemory()

                Me.Dispatcher.Invoke(New SetFolderNameTextCallback(AddressOf SetFoldersNameText), New Object() {"0 e-mails"})

                Me.Dispatcher.Invoke(New SetToolTipCallback(AddressOf SetToolTip), New Object() {"Done"})

                Me.Dispatcher.Invoke(New SetCursorCallback(AddressOf SetCursor), New Object() {Cursors.Arrow})

                SetProcessPriorities("End Review")

            End If

            Me.Dispatcher.Invoke(New FinalizeLoadCallback(AddressOf FinalizeLoad), New Object() {})

CleanExit:

        Catch ex As Exception
            MsgBox(ex.ToString)
        Finally
            ' common UI cleanup is already in FinalizeLoad
        End Try

        MemoryManagement.FlushMemory()

        'swOverall.Stop() : Console.WriteLine("Overall time to process: " & swOverall.ElapsedMilliseconds / 1000)

    End Sub

#Region "Load Folder Table"

    Sub FindAllFolders()

        If gCancelRefresh Then Exit Sub

        gScanningFolders = True

        Try

            gFolderButtonsOnOptionsWindowEnabled = False
            If gOptionsWindow IsNot Nothing Then
                gOptionsWindow.SafelyEnableOptionsFolderButtons()
            End If

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
            'MsgBox(sw.ElapsedMilliseconds.ToString)
            'Console.WriteLine(sw.ElapsedMilliseconds.ToString) : sw.Stop()

            gFolderTableIndex -= 1
            gFolderTableCurrentSize = gFolderTableIndex

            ReDim Preserve gFolderTable(gFolderTableIndex)
            ReDim gFolderNamesTable(gFolderTableIndex)
            ReDim gFolderNamesTableTrimmed(gFolderTableIndex)

            For i As Integer = 0 To gFolderTable.Length - 1
                gFolderNamesTable(i) = gFolderTable(i).FolderPath
                gFolderNamesTableTrimmed(i) = gFolderNamesTable(i).TrimStart("\")
            Next

            ' Detect special folders across all mailboxes
            gDeletedFolderIndex = -1
            gInboxFolderIndex = -1
            gSentFolderIndex = -1
            gInboxFolderIndices.Clear()
            gSentFolderIndices.Clear()
            gStoreDeleteFolders.Clear()

            ' First pass: locate Inbox/Sent and best delete folder (Deleted Items / Deleted / Trash) per store
            For i As Integer = 0 To gFolderTable.Length - 1

                Dim fInfo As FolderInfo = gFolderTable(i)
                If fInfo.DefaultItemType <> Microsoft.Office.Interop.Outlook.OlItemType.olMailItem Then
                    Continue For
                End If

                Dim nameUpper As String = System.IO.Path.GetFileName(fInfo.FolderPath).Trim().ToUpperInvariant()

                ' Track Inbox folders globally and per mailbox
                If nameUpper = "INBOX" Then
                    gInboxFolderIndices.Add(i)
                    If gInboxFolderIndex = -1 Then gInboxFolderIndex = i
                End If

                ' Track Sent folders globally and per mailbox
                If nameUpper = "SENT ITEMS" OrElse nameUpper = "SENT" Then
                    gSentFolderIndices.Add(i)
                    If gSentFolderIndex = -1 Then gSentFolderIndex = i
                End If

                ' Figure out a suitable delete folder for this store:
                ' Prefer Deleted Items, then Deleted, then Trash.
                Dim isDeletedItems As Boolean = (nameUpper = "DELETED ITEMS")
                Dim isDeleted As Boolean = (nameUpper = "DELETED")
                Dim isTrash As Boolean = (nameUpper = "TRASH")

                If Not (isDeletedItems OrElse isDeleted OrElse isTrash) Then
                    Continue For
                End If

                Dim storeId As String = fInfo.StoreID
                If String.IsNullOrEmpty(storeId) Then
                    Continue For
                End If

                Dim existing As StoreDeleteFolderInfo
                Dim hasExisting As Boolean = gStoreDeleteFolders.TryGetValue(storeId, existing)

                If Not hasExisting Then
                    ' First candidate for this store
                    gStoreDeleteFolders(storeId) = New StoreDeleteFolderInfo With {
                        .StoreId = storeId,
                        .FolderIndex = i
                    }
                Else
                    ' We may have a weaker candidate already; prefer Deleted Items > Deleted > Trash
                    Dim currentNameUpper As String = System.IO.Path.GetFileName(gFolderTable(existing.FolderIndex).FolderPath).Trim().ToUpperInvariant()

                    Dim currentScore As Integer = GetDeleteFolderPreferenceScore(currentNameUpper)
                    Dim newScore As Integer = GetDeleteFolderPreferenceScore(nameUpper)

                    If newScore > currentScore Then
                        existing.FolderIndex = i
                        gStoreDeleteFolders(storeId) = existing
                    End If
                End If

                ' Maintain backwards‑compatible global deleted index for legacy callers:
                ' pick the first 'Deleted Items' we see, then fall back to any previous.
                If gDeletedFolderIndex = -1 AndAlso (isDeletedItems OrElse isDeleted OrElse isTrash) Then
                    gDeletedFolderIndex = i
                End If

            Next

            ' Debug: ensure we found at least some inbox/sent folders
            ' MsgBox("Inboxes: " & gInboxFolderIndices.Count & " Sent: " & gSentFolderIndices.Count)

            gFolderButtonsOnOptionsWindowEnabled = True
            If gOptionsWindow IsNot Nothing Then
                gOptionsWindow.SafelyEnableOptionsFolderButtons()
            End If


            Dim ToolTipMessage As String = ""
            Dim ProgressBarMaxValue As Double

            If gRefreshAll OrElse gRefreshInbox OrElse gRefreshSent Then

                If gRefreshAll Then

                    ToolTipMessage = "E-mails from all included folders are being reviewed"

                    'ProcessBarMaxValue = 
                    ' 1 times the TotalEMails To Be Reviewed for processing all info but the trailer +
                    ' 2 times the TotalEMails To Be Reviewed for processing the trailer +
                    ' a time factor doing the recommendations
                    'ProgressBarMaxValue = (3 * lTotalEMailsToBeReviewed) + Int(lTotalEMailsToBeReviewed * (1 + My.Settings.RatioOfRecommendationToProcessingTime + 0.01))
                    ProgressBarMaxValue = lTotalEMailsToBeReviewed * 3
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

            gThresholdForReportingProgressOnTheProgressBar = Math.Max(50, lTotalEMailsToBeReviewed / 200) ' every 1/2 percent

            'gThresholdForReportingProgressOnTheProgressBar = Math.Max(50, lTotalEMailsToBeReviewed / 100) ' every percent


            Me.Dispatcher.Invoke(New SetToolTipCallback(AddressOf SetToolTip), New Object() {ToolTipMessage})
            Me.Dispatcher.Invoke(New SetProgressBarMaxValueCallback(AddressOf SetProgressBarMaxValue), New Object() {ProgressBarMaxValue})

        Catch ex As Exception

            MsgBox(ex.ToString)

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

        'Dim sStartFolderName As String = StartFolder.FolderPath.ToString
        'System.Diagnostics.Debug.WriteLine("Processing folder: " & sStartFolderName)

        Dim defaultItemType As Microsoft.Office.Interop.Outlook.OlItemType

        Try
            defaultItemType = StartFolder.DefaultItemType
        Catch ex As System.Runtime.InteropServices.COMException
            ' Skip folders that cannot be inspected due to Outlook/MAPI errors
            Exit Sub
        Catch
            ' Any other error getting DefaultItemType – skip this folder
            Exit Sub
        End Try

        ' Process the current folder only if it is a mail folder
        If defaultItemType = Microsoft.Office.Interop.Outlook.OlItemType.olMailItem Then
            Try
                AddAnEntry(StartFolder)
            Catch ex As System.Runtime.InteropServices.COMException
                ' Ignore folders that fail when adding an entry
            Catch
                ' Ignore and continue
            End Try
        End If

        ' Process all subfolders (using recursion), guarding against COM failures
        Dim subFolders As Microsoft.Office.Interop.Outlook.Folders
        Try
            subFolders = StartFolder.Folders
        Catch ex As System.Runtime.InteropServices.COMException
            ' Cannot enumerate subfolders for this folder
            Exit Sub
        Catch
            Exit Sub
        End Try

        If subFolders Is Nothing Then Exit Sub

        For Each oFolder As Microsoft.Office.Interop.Outlook.MAPIFolder In subFolders

            If gCancelRefresh Then Exit For

            Try
                AddFolder(oFolder)
            Catch ex As System.Runtime.InteropServices.COMException
                ' Skip any subfolder that errors
            Catch
                ' Ignore and continue with remaining subfolders
            End Try
        Next

    End Sub

    Private Sub AddAnEntry(ByRef Folder As Microsoft.Office.Interop.Outlook.MAPIFolder)

        ' Ensure the folder table is initialized at least once
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
        Dim folderNameUpper As String = System.IO.Path.GetFileName(CurrentFolderPath).Trim().ToUpperInvariant()
        Dim isInboxFolder As Boolean = (folderNameUpper = "INBOX")
        Dim isSentFolder As Boolean = (folderNameUpper = "SENT ITEMS" OrElse folderNameUpper = "SENT")

        If (gRefreshInbox AndAlso isInboxFolder) OrElse
       (gRefreshSent AndAlso isSentFolder) Then

            Include = True

        ElseIf gRefreshAll Then

            Include = Collection_of_folders_to_exclude_is_empty OrElse (Collection_of_folders_to_exclude.IndexOf(CurrentFolderPath) = -1)

        Else
            Include = False
        End If

        Dim folderItemCount As Integer = 0
        Try
            folderItemCount = Folder.Items.Count
        Catch
            folderItemCount = 0
        End Try

        If Include Then
            Me.Dispatcher.BeginInvoke(
            New SetFolderNameTextCallback(AddressOf SetFoldersNameText),
            New Object() {"Including " & CurrentFolderPath.TrimStart("\")})
            lTotalEMailsToBeReviewed += folderItemCount
        Else
            Me.Dispatcher.BeginInvoke(
            New SetFolderNameTextCallback(AddressOf SetFoldersNameText),
            New Object() {"Excluding " & CurrentFolderPath.TrimStart("\")})
        End If

        lTotalEMails += folderItemCount

    End Sub

#End Region

#Region "Load EMail Table"

    ' MD5-based fingerprint for trailers (restores original behaviour)
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

    Private Sub ProcessAllFolders()

        'Dim sw As New Stopwatch
        'sw.Start()

        '***************************************************************************
        'Step 1 initializations
        '***************************************************************************

        Dim iProgressBarValue As Double = 0

        lWhenSent = My.Settings.WhenSent

        EmailTableIndex = 0

        ReDim EmailTable(lTotalEMailsToBeReviewed)

        Dim strCollection As System.Collections.Specialized.StringCollection = New System.Collections.Specialized.StringCollection
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

        Dim ScanThisFolder As Boolean

        For x As Int16 = 0 To gFolderTableIndex

            If gCancelRefresh Then Exit For

            ScanThisFolder = False

            If gInboxFolderIndices.Contains(x) Then

                ScanThisFolder = gRefreshInbox AndAlso (gInboxFolderIndices.Count > 0)

            ElseIf gSentFolderIndices.Contains(x) Then

                ScanThisFolder = gRefreshSent AndAlso (gSentFolderIndices.Count > 0)

            ElseIf gRefreshAll Then

                ScanThisFolder = Collection_of_folders_to_exclude_is_empty OrElse (strCollection.IndexOf(gFolderNamesTable(x)) = -1)

            End If

            If ScanThisFolder Then
                If gFolderTable(x).DefaultItemType = Microsoft.Office.Interop.Outlook.OlItemType.olMailItem Then
                    Dim folder As Microsoft.Office.Interop.Outlook.MAPIFolder = Nothing
                    Try
                        folder = oNS.GetFolderFromID(gFolderTable(x).EntryID, gFolderTable(x).StoreID)
                    Catch
                        folder = Nothing
                    End Try

                    If folder IsNot Nothing Then
                        ProcessMailItems(x, folder, iProgressBarValue)
                        If gCancelRefresh Then Exit For
                    End If

                End If
            End If

        Next

        If gCancelRefresh Then GoTo EarlyExit

        strCollection = Nothing

        ReDim Preserve EmailTable(EmailTableIndex - 1)

        '***************************************************************************
        'Step 2 sort email table so that subjects are grouped together
        '***************************************************************************

        Dim lEMailTableSorter As New EMailTableSorter
        lEMailTableSorter.PrimarySortColumn = 1 ' subject
        lEMailTableSorter.PrimaryOrder = SortOrder.Ascending
        lEMailTableSorter.SecondarySortColumn = 4 ' date and time 
        lEMailTableSorter.SecondaryOrder = SortOrder.Descending
        Array.Sort(EmailTable, lEMailTableSorter)
        lEMailTableSorter = Nothing

        '***************************************************************************
        'Step 3 add trailers only when needed
        '***************************************************************************

        Dim addTrailerFlag As Boolean

        For x = 0 To EmailTable.Length - 1

            If gCancelRefresh Then Exit For

            Select Case x

                Case 0
                    If EmailTable.Length > 1 Then
                        addTrailerFlag = (EmailTable(x).sSubject = EmailTable(x + 1).sSubject)
                    Else
                        addTrailerFlag = False
                    End If

                Case < EmailTable.Length - 1
                    addTrailerFlag = (EmailTable(x).sSubject = EmailTable(x - 1).sSubject) OrElse
                                     (EmailTable(x).sSubject = EmailTable(x + 1).sSubject)

                Case Else
                    If EmailTable.Length > 1 Then
                        addTrailerFlag = (EmailTable(x).sSubject = EmailTable(x - 1).sSubject)
                    Else
                        addTrailerFlag = False
                    End If

            End Select

            If addTrailerFlag Then

                Dim body As String = EmailTable(x).sBody

                If String.IsNullOrEmpty(body) Then
                    EmailTable(x).sTrailer = ""
                Else
                    EmailTable(x).sTrailer = body
                End If

                EmailTable(x).sTrailer = EmailTable(x).sTrailer.Trim()

                If EmailTable(x).sTrailer.Length = 0 Then

                    ' Optional: previous attachment fallback logic if you still want it
                    ' (requires resolving the MailItem again or caching attachment info)
                    ' Otherwise, you can just use the Chr(255) marker:
                    EmailTable(x).sTrailer = Chr(255)

                Else

                    lLastIndex = EmailTable(x).sTrailer.LastIndexOf("Subject:")
                    If lLastIndex > -1 Then
                        EmailTable(x).sTrailer = EmailTable(x).sTrailer.Remove(0, lLastIndex + 8)
                    Else
                        lLastIndex = EmailTable(x).sTrailer.LastIndexOf("SUBJECT:")
                        If lLastIndex > -1 Then
                            EmailTable(x).sTrailer = EmailTable(x).sTrailer.Remove(0, lLastIndex + 8)
                        End If
                    End If

                    ' Remove stuff so the email chains can be properly linked together
                    EmailTable(x).sTrailer = EmailTable(x).sTrailer _
                        .Replace(" ", "") _
                        .Replace(vbCr, "") _
                        .Replace(vbLf, "") _
                        .Replace(">", "") _
                        .Replace(vbTab, "")

                    Dim lHoldSubject As String =
                        EmailTable(x).sSubject _
                            .Replace(" ", "") _
                            .Replace(vbCr, "") _
                            .Replace(vbLf, "") _
                            .Replace(">", "") _
                            .Replace(vbTab, "")

                    If lHoldSubject.Length > 0 Then
                        EmailTable(x).sTrailer = EmailTable(x).sTrailer.Replace(lHoldSubject, "")
                    End If

                    ' Only work with the first 240 chars to avoid endless growth at the end
                    If EmailTable(x).sTrailer.Length > 240 Then
                        EmailTable(x).sTrailer = EmailTable(x).sTrailer.Remove(240)
                    End If

                End If

                If EmailTable(x).sTrailer.Length > 16 Then
                    EmailTable(x).sTrailer = ComputeTrailerHash(EmailTable(x).sTrailer)
                End If

            End If

            lProgressBareRefreshingThresholdCounter += 2
            If lProgressBareRefreshingThresholdCounter > gThresholdForReportingProgressOnTheProgressBar Then
                iProgressBarValue += lProgressBareRefreshingThresholdCounter
                lProgressBareRefreshingThresholdCounter = 0
                Me.Dispatcher.Invoke(
                       New SetProgressBarValueCallback(AddressOf SetProgressBarValue),
                       New Object() {iProgressBarValue})
            End If

        Next

EarlyExit:

        'sw.Stop()
        'MsgBox(sw.ElapsedMilliseconds)

    End Sub

    'moved here for performance gains

    Private lBlankEMailDetailRecord As StructureOfEmailDetails
    Private lws As String
    Private lWhenSent As Boolean
    Private lLastIndex As Integer

    Const olFormatHTML = 2
    Const olFormatPlain = 1
    Const olFormatRichText = 3
    Const olFormatUnspecified = 0

    Private FlipFlop As Boolean = True

    Const WhenSentQuery As String = "SELECT EntryID, Subject, To, SenderEmailAddress, SentOn, UnRead from Folder Where MessageClass='IPM.Note'"
    Const WhenReceivedQuery As String = "SELECT EntryID, Subject, To, SenderEmailAddress, ReceivedTime, UnRead from Folder Where MessageClass='IPM.Note'"

    Private Query As String = "" ' moved here for efficiency

    Private Sub ProcessMailItems(ByVal originalFolder As Int16,
                             ByVal folder As Microsoft.Office.Interop.Outlook.MAPIFolder,
                             ByRef iProgressBarValue As Double)

        If gCancelRefresh Then Exit Sub

        Dim items As Microsoft.Office.Interop.Outlook.Items = Nothing

        Try
            items = folder.Items
        Catch
            Exit Sub
        End Try

        Dim sortField As String = If(lWhenSent, "[SentOn]", "[ReceivedTime]")
        Try
            items.Sort(sortField, True)
        Catch
        End Try

        ' Resolve mailbox / postbox name once per folder
        Dim mailboxName As String = GetMailboxNameFromFolderPath(folder.FolderPath, folder.StoreID)

        ' Track distinct mailbox names
        If Not String.IsNullOrEmpty(mailboxName) Then
            If Not _mailboxNames.Contains(mailboxName) Then
                _mailboxNames.Add(mailboxName)
                _mailboxCount = _mailboxNames.Count
            End If
        End If

        For i As Integer = 1 To items.Count

            If gCancelRefresh Then Exit For

            Dim obj As Object
            Try
                obj = items(i)
            Catch
                Continue For
            End Try

            Dim mail As Microsoft.Office.Interop.Outlook.MailItem =
            TryCast(obj, Microsoft.Office.Interop.Outlook.MailItem)
            If mail Is Nothing Then Continue For

            Dim emailDetail As StructureOfEmailDetails = lBlankEMailDetailRecord

            ' Resolve a friendly "From" address (gets around a quirk in Outlook / Exchange for messages coming from Exchange or certain connected accounts)
            Dim friendlyFrom As String = ""
            Try
                If mail.Sender IsNot Nothing Then
                    ' If this is already an SMTP address, use it directly
                    If String.Equals(mail.SenderEmailType, "SMTP", StringComparison.OrdinalIgnoreCase) Then
                        friendlyFrom = mail.SenderEmailAddress

                    Else
                        ' Try to resolve to an Exchange user and use PrimarySmtpAddress
                        Dim exUser As Microsoft.Office.Interop.Outlook.ExchangeUser =
                            TryCast(mail.Sender.GetExchangeUser(), Microsoft.Office.Interop.Outlook.ExchangeUser)

                        If exUser IsNot Nothing AndAlso
                           Not String.IsNullOrEmpty(exUser.PrimarySmtpAddress) Then

                            friendlyFrom = exUser.PrimarySmtpAddress

                        Else
                            ' Fallback – use whatever Outlook gives us
                            friendlyFrom = mail.SenderEmailAddress
                        End If
                    End If
                Else
                    friendlyFrom = mail.SenderEmailAddress
                End If
            Catch
                ' As a safety net, fall back to SenderEmailAddress on any COM error
                friendlyFrom = mail.SenderEmailAddress
            End Try

            With emailDetail

                .sOriginalFolderReferenceNumber = originalFolder
                .sOutlookEntryID = mail.EntryID
                .sSubject = CleanUpSubjectLine(If(mail.Subject, String.Empty))
                .sTo = If(mail.To, String.Empty)
                .sFrom = If(friendlyFrom, String.Empty)
                .sDateAndTime = If(lWhenSent, mail.SentOn, mail.ReceivedTime)
                .sUnRead = If(mail.UnRead,
                          System.Windows.FontWeights.Bold,
                          System.Windows.FontWeights.Normal)
                .sMailBoxName = mailboxName

                ' Only read Body when the subject isn't unique for this folder,
                ' or when a folder is large enough that chains are likely.
                ' For a quick win, you can start by always reading it, then optimize:
                .sBody = If(mail.Body, String.Empty)

            End With

            EmailTable(EmailTableIndex) = emailDetail
            EmailTableIndex += 1

            lProgressBareRefreshingThresholdCounter += 1
            If lProgressBareRefreshingThresholdCounter > gThresholdForReportingProgressOnTheProgressBar Then
                iProgressBarValue += lProgressBareRefreshingThresholdCounter
                lProgressBareRefreshingThresholdCounter = 0
                Me.Dispatcher.Invoke(
                New SetProgressBarValueCallback(AddressOf SetProgressBarValue),
                New Object() {iProgressBarValue})
            End If
        Next

    End Sub

    Private Function CleanUpSubjectLine(ByVal subjectline As String) As String

        'remove the FW: and RE: prefixes
        subjectline = subjectline.Trim()

        While subjectline.Length >= 3
            ' take at most first 3 characters
            Dim prefix As String = subjectline.Substring(0, 3).ToUpperInvariant()

            If prefix = "FW:" OrElse prefix = "RE:" Then
                ' strip prefix and trim again
                subjectline = subjectline.Substring(3).TrimStart()
            Else
                Exit While
            End If
        End While

        subjectline = subjectline.TrimStart()

        If subjectline.Length = 0 Then subjectline = " "

        Return subjectline

    End Function

#End Region

#Region "Establish and set rankings"

    Structure StructureOfEmailIDTable
        Friend EMailID As String
        Friend Count As Integer
    End Structure
    Structure StructureOfEmailAddressTableEntry
        Friend Entry() As StructureOfEmailIDTable
    End Structure
    Dim EmailIDFolderTable() As StructureOfEmailAddressTableEntry
    Dim EmailIDFolderTableIndex As Integer
    Private Sub EstablishRecomendations()

        If gRefreshAll Then
            'recommendations are only made when refresh all is selected
        Else
            Exit Sub
        End If

        Try

            'A second sort of the email table is required to subjects in order with thier trailers

            Dim lEMailTableSorter As New EMailTableSorter
            lEMailTableSorter.PrimarySortColumn = 1 ' subject
            lEMailTableSorter.PrimaryOrder = SortOrder.Ascending
            lEMailTableSorter.SecondarySortColumn = 4 ' date and time 
            lEMailTableSorter.SecondaryOrder = SortOrder.Descending
            Array.Sort(EmailTable, lEMailTableSorter)
            lEMailTableSorter = Nothing

            EstablishRatings_NumberOfUniqueEmailsInAFolder()

            EstablishRatings_Scoring()

        Catch ex As Exception

            MsgBox(ex.TargetSite.Name & " - " & ex.ToString)

        End Try

    End Sub

    Private Sub EstablishRatings_NumberOfUniqueEmailsInAFolder()

        'Set up for rating number of e-mails related to the same chain within a folder

        Try
            Dim MaximumNumberOfFoldersThatAUniqueEmailChainCanBeFoundWithin As Int16 = 500
            Dim CurrentSubjectAndTrailer As String = "|*| something unique |*|" & Chr(255)
            Dim PrevSubjectAndTrailer As String = ""
            Dim wi As Int16
            Dim MatchFound As Boolean

            ReDim UniqueSubjectsTable(lTotalEMailsToBeReviewed)
            UniqueSubjectsTableIndex = -1

            For x As Int32 = 0 To EmailTable.Length - 1

                PrevSubjectAndTrailer = CurrentSubjectAndTrailer
                CurrentSubjectAndTrailer = EmailTable(x).sSubject & EmailTable(x).sTrailer

                If CurrentSubjectAndTrailer <> PrevSubjectAndTrailer Then

                    UniqueSubjectsTableIndex += 1

                    ' Setup for a new entry
                    UniqueSubjectsTable(UniqueSubjectsTableIndex).SubjectAndTrailer = CurrentSubjectAndTrailer
                    UniqueSubjectsTable(UniqueSubjectsTableIndex).TotalOccurances = 1

                    ReDim UniqueSubjectsTable(UniqueSubjectsTableIndex).Occurances(MaximumNumberOfFoldersThatAUniqueEmailChainCanBeFoundWithin)


                    UniqueSubjectsTable(UniqueSubjectsTableIndex).Occurances(0).FolderReferenceCount = 1
                    UniqueSubjectsTable(UniqueSubjectsTableIndex).Occurances(0).FolderReference = EmailTable(x).sOriginalFolderReferenceNumber

                    If UniqueSubjectsTableIndex > 0 Then
                        For ii As Int16 = 0 To MaximumNumberOfFoldersThatAUniqueEmailChainCanBeFoundWithin - 1
                            If UniqueSubjectsTable(UniqueSubjectsTableIndex - 1).Occurances(ii).FolderReferenceCount = 0 Then
                                wi = UniqueSubjectsTable(UniqueSubjectsTableIndex - 1).TotalOccurances
                                ReDim Preserve UniqueSubjectsTable(UniqueSubjectsTableIndex - 1).Occurances(wi - 1)
                                Exit For
                            End If
                        Next
                    End If

                Else

                    ' where email is part of the same email chain
                    ' if first time found in a folder set folder counter to 1 
                    ' otherwise increment existing folder counter
                    MatchFound = False
                    For i As Int16 = 0 To UniqueSubjectsTable(UniqueSubjectsTableIndex).TotalOccurances - 1
                        If UniqueSubjectsTable(UniqueSubjectsTableIndex).Occurances(i).FolderReference = EmailTable(x).sOriginalFolderReferenceNumber Then
                            UniqueSubjectsTable(UniqueSubjectsTableIndex).Occurances(i).FolderReferenceCount += 1
                            MatchFound = True
                            Exit For
                        End If
                    Next

                    If Not MatchFound Then
                        wi = UniqueSubjectsTable(UniqueSubjectsTableIndex).TotalOccurances
                        UniqueSubjectsTable(UniqueSubjectsTableIndex).Occurances(wi).FolderReference = EmailTable(x).sOriginalFolderReferenceNumber
                        UniqueSubjectsTable(UniqueSubjectsTableIndex).Occurances(wi).FolderReferenceCount = 1
                        UniqueSubjectsTable(UniqueSubjectsTableIndex).TotalOccurances += 1
                    End If

                End If

            Next

            ReDim Preserve UniqueSubjectsTable(UniqueSubjectsTableIndex)

        Catch ex As Exception

            MsgBox(ex.TargetSite.Name & " - " & ex.ToString)

        End Try

    End Sub

    Private Sub EstablishRatings_Scoring()

        'For each unique email chain, rate the best folder to put it in
        '   1 point to each folder for each email in it that belongs to the same unique email chain

        Try

            Dim CurrentSubjectAndTrailer As String = "|*| something unique |*|" & Chr(255)
            Dim PrevSubjectAndTrailer As String = ""

            Dim CurrentSubjectShort As String = "|*| something unique |*|" & Chr(255)
            Dim PrevSubjectShort As String = ""

            Dim ii As Integer

            Dim FinalScoringTable(gFolderTable.Length - 1) As Integer

            Dim FolderNameUpperCase(gFolderTable.Length - 1) As String
            For i = 0 To gFolderTable.Length - 1
                FolderNameUpperCase(i) = gFolderNamesTable(i).ToUpper
            Next


            For i As Integer = 0 To EmailTable.Length - 1


                Dim ThisEMailsScore As Integer = 0

                PrevSubjectAndTrailer = CurrentSubjectAndTrailer
                CurrentSubjectAndTrailer = EmailTable(i).sSubject & EmailTable(i).sTrailer

                If CurrentSubjectAndTrailer <> PrevSubjectAndTrailer Then

                    For ii = ii To UniqueSubjectsTable.Length - 1  'start where we left off
                        If CurrentSubjectAndTrailer = UniqueSubjectsTable(ii).SubjectAndTrailer Then
                            Exit For
                        End If
                    Next

                    ' points to each folder that has an each email in it that belongs to the same unique email chain
                    Array.Clear(FinalScoringTable, 0, FinalScoringTable.Length)

                    For iScoreA As Integer = 0 To UniqueSubjectsTable(ii).TotalOccurances - 1
                        FinalScoringTable(UniqueSubjectsTable(ii).Occurances(iScoreA).FolderReference) =
                        UniqueSubjectsTable(ii).Occurances(iScoreA).FolderReferenceCount
                    Next
                    '***** Final Recommendations

                    ' Don't recommend the original folder
                    'FinalScoringTable(EmailTable(i).sOriginalFolderReferenceNumber) = 0

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

                    FindTheFolderWithTheGreatestScore(EmailTable(i).sRecommendedFolder1ReferenceNumber, FinalScoringTable)
                    FindTheFolderWithTheGreatestScore(EmailTable(i).sRecommendedFolder2ReferenceNumber, FinalScoringTable)
                    FindTheFolderWithTheGreatestScore(EmailTable(i).sRecommendedFolder3ReferenceNumber, FinalScoringTable)
                    EmailTable(i).sRecommendedFolderFinalReferenceNumber = EmailTable(i).sRecommendedFolder1ReferenceNumber

                Else

                    EmailTable(i).sRecommendedFolder1ReferenceNumber = EmailTable(i - 1).sRecommendedFolder1ReferenceNumber
                    EmailTable(i).sRecommendedFolder2ReferenceNumber = EmailTable(i - 1).sRecommendedFolder2ReferenceNumber
                    EmailTable(i).sRecommendedFolder3ReferenceNumber = EmailTable(i - 1).sRecommendedFolder3ReferenceNumber
                    EmailTable(i).sRecommendedFolderFinalReferenceNumber = EmailTable(i - 1).sRecommendedFolderFinalReferenceNumber

                End If

            Next

            UniqueSubjectsTable = Nothing

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub

    Private Sub FindTheFolderWithTheGreatestScore(ByRef ReferenceNumber As Integer, ByRef FinalScoringTable() As Integer)

        'The following code works with dotnet version 3.5 only, is replaced by the codeblock below it for use with dot net version 3.0
        'Find the entry with the greatest final score
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

            'If EmailTableIndex = 0 Then Exit Try

            ' When only Inbox/Sent are scanned, show those items directly
            If Not gRefreshAll AndAlso (gRefreshInbox OrElse gRefreshSent) Then

                ReDim gFinalRecommendationTable(EmailTableIndex - 1)

                Dim line As Integer = 0

                For x As Integer = 0 To EmailTableIndex - 1

                    Dim origFolder As Integer = EmailTable(x).sOriginalFolderReferenceNumber
                    Dim inInbox As Boolean = gInboxFolderIndices.Contains(origFolder)
                    Dim inSent As Boolean = gSentFolderIndices.Contains(origFolder)

                    ' respect ScanInbox/ScanSent toggles
                    If (gRefreshInbox AndAlso inInbox) OrElse (gRefreshSent AndAlso inSent) Then

                        Dim row As New ListViewRowClass
                        With EmailTable(x)
                            row.Index = line
                            row.MailBoxName = .sMailBoxName
                            row.Subject = .sSubject
                            row.Trailer = .sTrailer
                            row.From = .sFrom
                            row.xTo = .sTo
                            row.DateTime = .sDateAndTime
                            row.OriginalFolder = .sOriginalFolderReferenceNumber
                            row.RecommendedFolder1 = .sRecommendedFolder1ReferenceNumber
                            row.RecommendedFolder2 = .sRecommendedFolder2ReferenceNumber
                            row.RecommendedFolder3 = .sRecommendedFolder3ReferenceNumber
                            row.RecommendedFolderFinal = .sRecommendedFolderFinalReferenceNumber
                            row.OutlookEntryID = .sOutlookEntryID
                            row.UnRead = .sUnRead
                        End With

                        gFinalRecommendationTable(line) = row
                        line += 1
                    End If
                Next

                If line = 0 Then
                    ' nothing to show
                    ReDim gFinalRecommendationTable(0)
                    Me.Dispatcher.Invoke(New SetListViewItemCallback(AddressOf SetListViewItem),
                                     New Object() {gFinalRecommendationTable})
                    lTotalRecommendations = 0
                    Exit Try
                End If

                ReDim Preserve gFinalRecommendationTable(line - 1)

                ' sort by Subject (same as normal path)
                Dim sorter As New FinalRecommendationTableSorter
                sorter.PrimaryColumnToSort = "MailBoxName"
                sorter.SortOrder = FinalRecommendationTableSorter.MySortOrder.Ascending
                Array.Sort(gFinalRecommendationTable, sorter)
                sorter = Nothing

                Me.Dispatcher.Invoke(New SetListViewItemCallback(AddressOf SetListViewItem),
                                 New Object() {gFinalRecommendationTable})

                lTotalRecommendations = line

                Exit Try

            End If

            If EmailTableIndex = 0 Then Exit Try

            ReDim gFinalRecommendationTable(EmailTableIndex)

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

            For x As Integer = 0 To EmailTableIndex - 1

                lRecommendedIndexForAllEntriesInChainFinal = -1
                lRecommendedIndexForAllEntriesInChain1 = -1
                lRecommendedIndexForAllEntriesInChain2 = -1
                lRecommendedIndexForAllEntriesInChain3 = -1

                lFlagThisEmailChain = False

                lNextIndex = x
                lFirstSubjectPlusTrailer = EmailTable(x).sSubject & EmailTable(x).sTrailer
                lNextSubjectPlusTrailer = lFirstSubjectPlusTrailer

                ' for each email chain
                ' flag the chain to be reported if it contains an inbox item, sent item, or and email store anyplace other than the recommended folder
                While (lFirstSubjectPlusTrailer = lNextSubjectPlusTrailer) And (lNextIndex <= (EmailTableIndex - 1))

                    If gInboxFolderIndices.Contains(EmailTable(lNextIndex).sOriginalFolderReferenceNumber) Then

                        If gRefreshInbox Then lFlagThisEmailChain = True

                    ElseIf gSentFolderIndices.Contains(EmailTable(lNextIndex).sOriginalFolderReferenceNumber) Then

                        If gRefreshSent Then lFlagThisEmailChain = True

                    Else

                        If EmailTable(lNextIndex).sRecommendedFolderFinalReferenceNumber > -1 Then

                            If EmailTable(lNextIndex).sOriginalFolderReferenceNumber <> EmailTable(lNextIndex).sRecommendedFolderFinalReferenceNumber Then

                                lRecommendedIndexForAllEntriesInChainFinal = EmailTable(lNextIndex).sRecommendedFolderFinalReferenceNumber
                                lRecommendedIndexForAllEntriesInChain1 = EmailTable(lNextIndex).sRecommendedFolder1ReferenceNumber
                                lRecommendedIndexForAllEntriesInChain2 = EmailTable(lNextIndex).sRecommendedFolder2ReferenceNumber
                                lRecommendedIndexForAllEntriesInChain3 = EmailTable(lNextIndex).sRecommendedFolder3ReferenceNumber
                                lFlagThisEmailChain = True
                                Exit While

                            End If
                        End If

                    End If

                    lNextIndex += 1
                    If (lNextIndex <= (EmailTableIndex - 1)) Then
                        lNextSubjectPlusTrailer = EmailTable(lNextIndex).sSubject & EmailTable(lNextIndex).sTrailer
                    End If

                End While

                'ensure if an e-mail chain is flagged the a recommendation is made if at all possible
                'the following covers the case where there are inbox or sent items and all filed emails are in the smame folder
                If lFlagThisEmailChain Then

                    If lRecommendedIndexForAllEntriesInChainFinal = -1 Then

                        lNextIndex = x
                        lFirstSubjectPlusTrailer = EmailTable(x).sSubject & EmailTable(x).sTrailer
                        lNextSubjectPlusTrailer = lFirstSubjectPlusTrailer

                        While (lFirstSubjectPlusTrailer = lNextSubjectPlusTrailer) And (lNextIndex <= (EmailTableIndex - 1))

                            If Not gInboxFolderIndices.Contains(EmailTable(lNextIndex).sOriginalFolderReferenceNumber) Then

                                If Not gSentFolderIndices.Contains(EmailTable(lNextIndex).sOriginalFolderReferenceNumber) Then

                                    lRecommendedIndexForAllEntriesInChainFinal = EmailTable(lNextIndex).sRecommendedFolderFinalReferenceNumber
                                    lRecommendedIndexForAllEntriesInChain1 = EmailTable(lNextIndex).sRecommendedFolder1ReferenceNumber
                                    lRecommendedIndexForAllEntriesInChain2 = EmailTable(lNextIndex).sRecommendedFolder2ReferenceNumber
                                    lRecommendedIndexForAllEntriesInChain3 = EmailTable(lNextIndex).sRecommendedFolder3ReferenceNumber
                                    lFlagThisEmailChain = True
                                    Exit While

                                End If

                            End If

                            lNextIndex += 1

                            If (lNextIndex <= (EmailTableIndex - 1)) Then
                                lNextSubjectPlusTrailer = EmailTable(lNextIndex).sSubject & EmailTable(lNextIndex).sTrailer
                            End If


                        End While

                    End If

                End If


                If lFlagThisEmailChain Then

                    Dim lStartingSubjectPlusTrailer As String = EmailTable(x).sSubject & EmailTable(x).sTrailer
                    Dim lChainEntry As Integer = x

                    While lStartingSubjectPlusTrailer = EmailTable(lChainEntry).sSubject & EmailTable(lChainEntry).sTrailer

                        With EmailTable(lChainEntry)

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
                        If lChainEntry > EmailTableIndex - 1 Then
                            Exit While
                        End If

                    End While

                    x = lChainEntry - 1

                End If

            Next

            EmailTable = Nothing

            ReDim Preserve gFinalRecommendationTable(lLineNumber - 1)

            Dim lFinalRecommendationTableSorter As New FinalRecommendationTableSorter
            lFinalRecommendationTableSorter.PrimaryColumnToSort = "MailBoxName"
            lFinalRecommendationTableSorter.SortOrder = FinalRecommendationTableSorter.MySortOrder.Ascending
            Array.Sort(gFinalRecommendationTable, lFinalRecommendationTableSorter)
            lFinalRecommendationTableSorter = Nothing

            Me.Dispatcher.Invoke(New SetListViewItemCallback(AddressOf SetListViewItem), New Object() {gFinalRecommendationTable})

            lTotalRecommendations = lLineNumber

        Catch ex As Exception

            MsgBox(ex.TargetSite.Name & " - " & ex.ToString)

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
        'ignor doubleclick if it happened within the bounds of the scrollbar

        If ListView1.ActualWidth - e.GetPosition(Me.ListView1).X > 15 Then
            UpdateDetails()
            RemoveHandler ListView1.SelectionChanged, AddressOf ListView1_SelectionChanged
            SelectAllMembersOfAnEmailChain()
            AddHandler ListView1.SelectionChanged, AddressOf ListView1_SelectionChanged
        End If

    End Sub

    Private Sub MainWindow_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseDown
        MenuKeyStrokeOverRide = False
    End Sub

    Private Sub ListView1_MouseEnter(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles ListView1.MouseEnter

        If Me.Cursor IsNot Cursors.Wait Then
            Me.Cursor = Cursors.Hand
        End If

    End Sub

    Private Sub ListView1_MouseLeave(ByVal sender As Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles ListView1.MouseLeave

        If Me.Cursor IsNot Cursors.Wait Then
            Me.Cursor = Cursors.Arrow
        End If

    End Sub

    Private Sub ListView1_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ListView1.SelectionChanged

        If gSupressUpdatesToDetailBox Then Exit Sub

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
            Dim selected = TryCast(ListView1.SelectedItem, ListViewRowClass)
            If selected Is Nothing Then
                BlankOutDetails()
                Exit Try
            End If

            With selected
                BlankOutDetails()

                Me.tbDetailSubject.Text = .Subject
                Me.tbDetailFrom.Text = .From
                Me.tbDetailTo.Text = .xTo
                Me.tbDetailDateTime.Text = Format(.DateTime, gPreferredDateFormat & " " & gPreferredTimeFormat)
                Me.tbDetailOrginal.Text = gFolderNamesTable(.OriginalFolder).TrimStart("\"c)

                ' Only offer pick folders when:
                ' 1) Scan Filed Emails is enabled (gRefreshAll = True), AND
                ' 2) There is a valid recommended folder (RecommendedFolderFinal >= 0)
                Dim hasValidRecommendation As Boolean = (.RecommendedFolderFinal >= 0)
                Dim shouldOfferPicks As Boolean = gRefreshAll AndAlso hasValidRecommendation

                If shouldOfferPicks Then
                    Me.tbDetailTarget1.Text = gFolderNamesTable(.RecommendedFolderFinal)
                Else
                    Me.tbDetailTarget1.Text = ""
                End If

                If gPickAFolderWindow IsNot Nothing Then
                    If shouldOfferPicks Then
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
            ' consider logging ex
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
    MenuUndo.Click, MenuHelpSub.Click, MenuAbout.Click, MenuOptions.Click, MenuRefresh.Click,
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

    Private Sub ToggleReadStateForSelectedItem()

        Try

            Dim selectedRow As ListViewRowClass = TryCast(ListView1.SelectedItem, ListViewRowClass)
            If selectedRow Is Nothing Then
                Exit Sub
            End If

            ' Ensure Outlook is running before trying to update an item
            If Not IsOutlookIsRunning() Then
                EnsureOutlookIsRunning()
            End If

            If oNS Is Nothing Then
                Exit Sub
            End If

            ' Get the Outlook item
            Dim entryId As String = selectedRow.OutlookEntryID
            If String.IsNullOrEmpty(entryId) Then
                Exit Sub
            End If

            Dim mailItem As Microsoft.Office.Interop.Outlook.MailItem = Nothing

            Try
                ' We do not have to pass StoreID here – behaviour matches existing OpenAnEmail
                mailItem = TryCast(oNS.GetItemFromID(entryId), Microsoft.Office.Interop.Outlook.MailItem)
            Catch
                mailItem = Nothing
            End Try

            If mailItem Is Nothing Then
                Exit Sub
            End If

            ' Toggle the unread flag in Outlook
            Dim currentlyUnread As Boolean = mailItem.UnRead
            mailItem.UnRead = Not currentlyUnread
            mailItem.Save()

            ' Reflect the new state in the ListView row
            Dim index As Integer = ListView1.SelectedIndex
            If index < 0 Then
                Exit Sub
            End If

            Dim updatedRow As ListViewRowClass = CType(ListView1.Items(index), ListViewRowClass)

            If mailItem.UnRead Then
                updatedRow.UnRead = System.Windows.FontWeights.Bold
            Else
                updatedRow.UnRead = System.Windows.FontWeights.Normal
            End If

            ' Replace item in the ListView so the binding updates
            ListView1.Items.RemoveAt(index)
            ListView1.Items.Insert(index, updatedRow)
            ListView1.SelectedIndex = index

            ' Also update details pane and counters
            UpdateDetails()

        Catch ex As Exception
            MsgBox(ex.Message,
                   MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly,
                   "FileFriendly - Toggle Read/Unread Failed")
        End Try

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
                    gRefreshAll = My.Settings.ScanAll
                    gAutoChainSelect = My.Settings.AutoChainSelect

                    If gARefreshIsRequired Then
                        RefreshGrid()
                    End If

                Case Is = "Undo"

                    If My.Settings.ConfirmUndo Then
                        If ConfirmActionMessage(Action) Then UndoLastAction()
                    Else
                        UndoLastAction()
                    End If

                Case Is = "Refresh"

                    ' If we are currently refreshing, this acts as Stop Refresh
                    If gIsRefreshing Then
                        gCancelRefresh = True
                        ' UI will be reset when the background task notices and exits
                        Exit Select
                    End If

                    If ActionLogIndex > 0 Then
                        If ShowMessageBox("FileFriendly",
                                          CustomDialog.CustomDialogIcons.Question,
                                          "Please note:",
                                          "If you refesh you will no longer be able to undo the changes you have made up until now." & vbCrLf & vbCrLf &
                                          "Would you still like to refresh?",
                                          "You will however be able to undo future changes.",
                                          "",
                                          CustomDialog.CustomDialogIcons.None,
                                          CustomDialog.CustomDialogButtons.YesNo,
                                          CustomDialog.CustomDialogResults.Yes) = CustomDialog.CustomDialogResults.No Then
                            Exit Select
                        End If
                    End If

                    MenuOptionEnabled("Undo", False)
                    ActionLogIndex = 0

                    gPickARefreshModeWindow = New PickARefreshMode
                    gPickARefreshModeWindow.ShowDialog()
                    gPickARefreshModeWindow = Nothing

                    If gRefreshConfirmed Then

                        If gRefreshInbox Or gRefreshSent Or gRefreshAll Then
                            MenuRefresh.Foreground = gForegroundColourEnabled
                            MenuActions.Foreground = gForegroundColourEnabled
                            ' start new refresh
                            RefreshGrid()
                        Else
                            ShowMessageBox("FileFriendly",
                                           CustomDialog.CustomDialogIcons.Warning,
                                           "Note!",
                                           "Inbox, sent items and other folders shouldn`t all be unchecked at the same time.",
                                           "If you uncheck all three then there will be nothing to review!",
                                           "",
                                           CustomDialog.CustomDialogIcons.None,
                                           CustomDialog.CustomDialogButtons.OK,
                                           CustomDialog.CustomDialogResults.OK)
                            ClearGrid()
                        End If

                    End If

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

                Case Is = "ViewRecommendedFolder"

                    If flag Then
                        Me.Label7.Visibility = Windows.Visibility.Visible
                        Me.TabControl2.Height = gOriginalTabControl2Height
                    Else
                        Me.Label7.Visibility = Windows.Visibility.Hidden
                        Me.Row3.Height = New System.Windows.GridLength(Me.Row3.ActualHeight - 20, GridUnitType.Auto)
                        Me.TabControl2.Height = Me.TabControl2.ActualHeight - 20
                    End If

                Case Is = "ViewFolderWindow"
                    ShowFolderWindow()

                Case Is = "Help"
                    System.Diagnostics.Process.Start(gHelpWebPage)
                    System.Threading.Thread.Sleep(3000)

                Case Is = "About"

                    gAboutWindow = New LicenseWindow
                    gAboutWindow.ShowDialog()
                    gAboutWindow = Nothing

            End Select

        Catch ex As Exception

            MsgBox(ex.ToString)

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

            ShowMessageBox("FileFriendly",
                           CustomDialog.CustomDialogIcons.Warning,
                           "Note!",
                           "Inbox, sent items and other folders shouldn`t all be unchecked at the same time.",
                           "If you uncheck all three then there will be nothing to review!",
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

            ShowMessageBox("FileFriendly",
                   CustomDialog.CustomDialogIcons.Warning,
                   "Note!",
                   "Read and Unread shouldn`t both be unchecked at the same time.",
                   "If you uncheck them both then there will be nothing to review!",
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

        Me.Close()

    End Sub

    Friend Declare Function GetDesktopWindow Lib "user32" () As IntPtr

    Private Sub OpenAnEmail()

        Try

            If ListView1.SelectedItem Is Nothing OrElse
               ListView1.SelectedItem.OutlookEntryID Is Nothing OrElse
               ListView1.SelectedItem.OutlookEntryID.Length = 0 Then
                Exit Try
            End If

            ' Ensure Outlook is running before trying to display an item
            If Not IsOutlookIsRunning() Then
                EnsureOutlookIsRunning()
            End If

            Dim EntryID As String = ListView1.SelectedItem.OutlookEntryID

            ' Open with Outlook object model only
            oNS.GetItemFromID(EntryID).Display()

            'if the email isn't already marked then mark it unread
            Dim index As Integer = ListView1.SelectedIndex
            If ListView1.Items(index).Unread = System.Windows.FontWeights.Bold Then
                Dim hold As ListViewRowClass = ListView1.SelectedItem
                hold.UnRead = System.Windows.FontWeights.Normal
                ListView1.Items.RemoveAt(index)
                ListView1.Items.Insert(index, hold)
                ListView1.SelectedIndex = index
                UpdateDetails()
            End If

        Catch ex As Exception
            MsgBox(ex.Message & vbCrLf & vbCrLf & "If Outlook is not running please start it and try again.", MsgBoxStyle.Exclamation, "FileFriendly - Open Fail")
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

            'point back to the last populated log entry
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
                ' effectively undeletes the list view item once the list view is refreshed
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

            'determine which entry should be the new selected entry 
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
                Else
                    Me.MenuRefresh.Foreground = gForegroundColourDisabled
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

                If gRefreshAll Then
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

        Dim ReturnCode As String

        Try
            ' Ensure session is valid
            If oNS Is Nothing OrElse oApp Is Nothing Then
                Throw New InvalidOperationException("Outlook session is not initialized.")
            End If

            If Not IsOutlookIsRunning() Then
                Throw New InvalidOperationException("Microsoft Outlook is not running.")
            End If

            oMailItem = oNS.GetItemFromID(EmailID, SourceStoreID)
            oTargetFolder = oNS.GetFolderFromID(TargetEntryID, TargetStoreID)

            If oMailItem Is Nothing OrElse oTargetFolder Is Nothing Then
                Throw New InvalidOperationException("Unable to resolve the selected e-mail or target folder in Outlook.")
            End If

            'Do the move
            Dim oMovedEmail As Microsoft.Office.Interop.Outlook.MailItem = oMailItem.Move(oTargetFolder)

            'Get new Entry ID
            Dim MovedEntryID As String = oMovedEmail.EntryID

            'unless it was the result of an undo request, record the action
            If Action <> "Undo" Then
                LogAction(Action, FixedIndex, MovedEntryID, TargetStoreID, SourceEntryID, SourceStoreID)
            End If

            ReturnCode = MovedEntryID

            oMovedEmail = Nothing

        Catch comEx As System.Runtime.InteropServices.COMException
            MsgBox("FileFriendly could not complete the requested action in Outlook." & vbCrLf & vbCrLf &
               "Details: " & comEx.Message,
               MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly,
               "FileFriendly - Outlook Error")
            ReturnCode = ""
        Catch ex As Exception
            MsgBox(ex.Message,
               MsgBoxStyle.Exclamation Or MsgBoxStyle.OkOnly,
               "FileFriendly - Outlook Error")
            ReturnCode = ""
        End Try

        Return ReturnCode

    End Function

    Private Sub LogAction(ByVal Action As String, ByVal FixedIndex As Integer, Optional ByVal MovedEntryID As String = "", Optional ByVal TargetStoreID As String = "", Optional ByVal SourceEntryID As String = "", Optional ByVal SourceStoreID As String = "")

        ' Record what happened (so that it can be undone later if necissary)

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
            ' MsgBox(ex.ToString) suppress error message here
        End Try

    End Sub

    Private SelectedListViewItem As New ListViewRowClass

    Private Sub ActionRequestAgainstAllSelectedItems(ByVal Action As String, ByRef ListView1 As ListView)

        If ListView1.SelectedItems.Count = 0 Then Exit Sub

        Me.ForceCursor = True
        Me.Cursor = Cursors.Wait

        gSupressUpdatesToDetailBox = True

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

        gSupressUpdatesToDetailBox = False

        UpdateDetails()

        Me.Cursor = Cursors.None

        Try

            ' the following works around a bug, where the entire list view is selected above the selected entry
            ' following an action and when the shift key is held down
            If ListView1.SelectedItems.Count = 1 Then

                RemoveHandler ListView1.SelectionChanged, AddressOf ListView1_SelectionChanged

                SelectedListViewItem = ListView1.SelectedItems(0)
                ListView1.SelectedItems.Clear()
                ListView1.SelectionMode = SelectionMode.Single
                ListView1.SelectedItem = SelectedListViewItem
                ListView1.SelectionMode = SelectionMode.Extended

                AddHandler ListView1.SelectionChanged, AddressOf ListView1_SelectionChanged


            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Function GetDeleteFolderIndexForStore(ByVal storeId As String) As Integer

        If Not String.IsNullOrEmpty(storeId) Then
            Dim info As StoreDeleteFolderInfo
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
            ' Ensure Outlook is running for any action that moves/deletes messages
            If Action = "File" OrElse Action = "Delete" Then
                If Not IsOutlookIsRunning() Then
                    ' Try to start Outlook and refresh the MAPI session
                    EnsureOutlookIsRunning()

                    ' Rebind oApp and oNS in case the old session is invalid
                    Try
                        oApp = CType(CreateObject("Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
                        oNS = oApp.GetNamespace("MAPI")
                    Catch exRebind As Exception
                        MsgBox("FileFriendly cannot access Microsoft Outlook." & vbCrLf & vbCrLf &
                           "The requested action cannot be completed until Outlook is available.",
                           MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly,
                           "FileFriendly - Outlook Not Available")
                        Exit Sub
                    End Try
                End If
            End If

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

                                Dim EmailID As String = ListView1.Items(IndexToAction).OutlookEntryID

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
                                                  EmailID,
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
                     CustomDialog.CustomDialogIcons.Stop,
                     "Opps",
                     "You've performed " & ActionLogMaxEntries & " actions, and that's exactly the limit I can remember!",
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

            'determine which entry should be the new selected entry 
            Dim IndexToBePositionedAt As Integer = 0
            If ListView1.Items.Count > 0 Then
                For i As Integer = IndexToAction - 1 To 0 Step -1
                    If ListView1.Items(i).subject IsNot Nothing Then
                        IndexToBePositionedAt = i + 1
                        Exit For
                    End If
                Next
            End If

            ReindexListView(ListView1)

            ' set focus on next listview item
            If IndexToBePositionedAt > ListView1.Items.Count - 1 Then
                IndexToBePositionedAt = ListView1.Items.Count - 1
            End If

            If ListView1.Items.Count > 0 Then
                If IndexToBePositionedAt >= 0 Then

                    If ListView1.Items(IndexToBePositionedAt).subject = Nothing Then
                        IndexToBePositionedAt += 1
                        If IndexToBePositionedAt > ListView1.Items.Count - 1 Then
                            IndexToBePositionedAt = ListView1.Items.Count - 1
                        End If
                    End If

                    ListView1.SelectedItem = ListView1.Items(IndexToBePositionedAt)

                    If gAutoChainSelect Then SelectAllMembersOfAnEmailChain()

                    ListView1.Focus()
                End If
            End If

        Catch ex As Exception

            MsgBox(ex.TargetSite.Name & " - " & ex.ToString)

        End Try

    End Sub

    Private Sub ReindexListView(ByRef lv As ListView)

        For x As Integer = 0 To lv.Items.Count - 1
            lv.Items(x).index = x
        Next

    End Sub

    Private Sub RemoveAnEntry(ByVal IndexToAction As Integer)

        gFinalRecommendationTable(ListView1.Items(IndexToAction).FixedIndex).Index = -1
        ListView1.Items.RemoveAt(IndexToAction)

    End Sub

#Region "Ensure Outlook is running if needed"

    Private Sub EnsureOutlookIsRunning()

        Dim mos As ManagementObjectSearcher = Nothing
        Dim mo As ManagementObject = Nothing
        Try

            mos = New ManagementObjectSearcher("SELECT NAME FROM Win32_Process")
            For Each mo In mos.Get
                If UCase(mo("Name").ToString) = "OUTLOOK.EXE" Then
                    ' Outlook is running
                    Exit Try
                End If
            Next

            BeginProgramNow(GetOutlookPathName)

            For x As Integer = 1 To 15
                Thread.Sleep(1000)
                mos = New ManagementObjectSearcher("SELECT NAME FROM Win32_Process")
                For Each mo In mos.Get
                    If UCase(mo("Name").ToString) = "OUTLOOK.EXE" Then
                        Thread.Sleep(5000)
                        Me.BringIntoView()
                        Exit Try
                    End If
                Next
            Next
            MsgBox("Outlook failed to start")

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        mo.Dispose()
        mos.Dispose()

    End Sub

    Friend Function IsOutlookIsRunning() As Boolean

        Dim ReturnCode As Boolean = False

        Dim mos As ManagementObjectSearcher = Nothing
        Dim mo As ManagementObject = Nothing
        Try
            mos = New ManagementObjectSearcher("SELECT NAME FROM Win32_Process")
            For Each mo In mos.Get
                If UCase(mo("Name").ToString) = "OUTLOOK.EXE" Then
                    ReturnCode = True
                    Exit Try
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        mo.Dispose()
        mos.Dispose()

        Return ReturnCode

    End Function

    Friend Function GetOutlookPathName() As String

        Dim ReturnCode As String = ""

        'Commented out code was moved outside subroutine for performance
        'Dim oApp As Microsoft.Office.Interop.Outlook.Application
        'Dim oNS As Microsoft.Office.Interop.Outlook.NameSpace
        Dim MyFolder As Microsoft.Office.Interop.Outlook.MAPIFolder

        'oApp = CreateObject("Outlook.Application")
        'oNS = oApp.GetNamespace("MAPI")
        MyFolder = oNS.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox)

        Dim mos As ManagementObjectSearcher
        Dim mo As ManagementObject

        Try
            mos = New ManagementObjectSearcher("SELECT Name, ExecutablePath FROM Win32_Process")
            For Each mo In mos.Get
                If UCase(mo("Name").ToString) = "OUTLOOK.EXE" Then
                    ReturnCode = mo("ExecutablePath").ToString.Trim
                    Exit For
                End If
            Next
        Catch ex As Exception
        End Try

        Return ReturnCode

    End Function


    Friend Sub BeginProgramNow(ByVal sProgram As String, Optional ByVal sArguments As String = "")

        Try

            Dim ws As String = "Starting a new program: " & vbCrLf & sProgram
            If sArguments.Length > 0 Then ws &= vbCrLf & sArguments

            sProgram = sProgram.Trim
            Dim TestProgramName As String = " " & UCase(sProgram)
            sArguments = sArguments.Trim

            Dim myProcess As New Process
            myProcess.StartInfo.Verb = "open"
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.UseShellExecute = True
            If sArguments.Length > 0 Then myProcess.StartInfo.Arguments = sArguments
            myProcess.StartInfo.FileName = sProgram
            myProcess.Start()

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

        If _mailboxCount <= 1 Then
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
        Dim timeHeader As String = TryCast(timeColumn.Header, String)

        Dim minFixedWidth As Double = 18

        Dim chain1Width As Double = Math.Max(minFixedWidth, measureText(chain1Header, baseTypeface) + headerPadding)
        Dim chain2Width As Double = Math.Max(minFixedWidth, measureText(chain2Header, baseTypeface) + headerPadding)

        Dim dateWidth As Double = Math.Max(80, measureText(If(dateHeader, "Date"), baseTypeface) + headerPadding)

        ' Time width based on the widest possible time string for the current format,
        ' measured in both normal and bold, then taking the larger.
        Dim widestTimeSample As Date = New Date(2000, 12, 31, 23, 59, 59)
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

        Dim mailboxVisible As Boolean = (_mailboxCount > 1)
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

        'restore to last postion
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

    Private Sub ListViewColumnHeaderClickedHandler(ByVal sender As Object, ByVal e As RoutedEventArgs)

        Me.Dispatcher.Invoke(New SetCursorCallback(AddressOf SetCursor), New Object() {Cursors.Wait})

        Try

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

                    Dim lFinalRecommendationTableSorter As New FinalRecommendationTableSorter
                    lFinalRecommendationTableSorter.PrimaryColumnToSort = header
                    lFinalRecommendationTableSorter.SortOrder = direction
                    Array.Sort(gFinalRecommendationTable, lFinalRecommendationTableSorter)
                    lFinalRecommendationTableSorter = Nothing

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

                End If

            End If

            ApplyFilter()

        Catch ex As Exception

            MsgBox(ex.ToString)
        End Try

        Me.Dispatcher.Invoke(New SetCursorCallback(AddressOf SetCursor), New Object() {Cursors.Arrow})

    End Sub


End Class

Public Class EMailTableSorter

    Implements System.Collections.IComparer

    Public Enum SortOrder As Integer
        None = 0
        Ascending = 1
        Descending = 2
    End Enum

    Private ObjectCompare As Comparer 'CaseInsensitiveComparer

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

    Private ObjectCompare As Comparer 'CaseInsensitiveComparer

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
