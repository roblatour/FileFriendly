Module modSharedDataAndRoutines

    ' Shared Data

    Friend ReadOnly gGithubWebPage As String = "https://github.com/roblatour/FileFriendly"
    Friend gHelpWebPage As String = "https://github.com/roblatour/FileFriendly/blob/main/Help/" ' the help filename be appended at startup

    Private ReadOnly gCurrentVersionWebPage As String = "https://raw.githubusercontent.com/roblatour/FileFriendly/refs/heads/main/versionControl/currentversion.txt"
    Private ReadOnly gDownloadWebPage = "https://github.com/roblatour/FileFriendly/releases/latest"

    Friend gAppVersion As Version
    Friend gAppVersionString As String

    Friend gScanningFolders As Boolean = True
    Friend gARefreshIsRequired As Boolean = False
    Friend gMinimizeMaximizeAllowed As Boolean = False

    Friend gDockSound(1) As Byte

    Friend gMainWindowIsMaximized As Boolean = False
    Friend gFolderButtonsOnOptionsWindowEnabled As Boolean = False

    Friend gAutoChainSelect As Boolean = False

    Friend gmwHeight As Double
    Friend gmwWidth As Double
    Friend gmwTop As Double
    Friend gmwLeft As Double
    Friend gOverridePickAWindowHeight As Boolean = False

    Friend gMinimizedAtEarlyStartup As Boolean = False

    Friend gRegisteredTo As String = ""
    Friend gFreeUpgrades As String = ""

    Friend gMostRecentDateInOutlook As DateTime = Now.AddMonths(-60)

    Friend gWindowDocked As Boolean = True
    Friend gPickAFolderWindowWasDocedWhenMainWindowWasMaximimized As Boolean = True
    Friend gSentText As String

    Friend gCurrentSortOrder As String = "Subject"

    Friend gFolderTableIndex As Integer = 0
    Friend Const gFolderTableIncrement As Integer = 1000
    Friend gFolderTableCurrentSize As Integer = 1000

    ' Holds Outlook folder identity without keeping a COM object across threads
    Friend Structure FolderInfo
        Friend EntryID As String
        Friend StoreID As String
        Friend FolderPath As String
        Friend DefaultItemType As Microsoft.Office.Interop.Outlook.OlItemType
    End Structure

    ' Outlook folder table now stores FolderInfo instead of MAPIFolder
    Friend gFolderTable() As FolderInfo
    Friend gFolderNamesTable(gFolderTableCurrentSize) As String
    Friend gFolderNamesTableTrimmed(gFolderTableCurrentSize) As String

    Friend gPreferredDateFormat As String = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern
    Friend gPreferredTimeFormat As String = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.LongTimePattern

    'Shared Windows
    Public gMainWindow As MainWindow
    Public gPickAFolderWindow As PickAFolder
    Public gOptionsWindow As OptionsWindow
    Public gFolderReviewWindow As FolderReviewWindow
    Friend gPickARefreshModeWindow As PickARefreshMode
    Friend gAboutWindow As LicenseWindow

    Friend gRefreshInbox, gRefreshSent, gRefreshAll, gRefreshConfirmed As Boolean
    Friend gBypassRefreshPrompt As Boolean = False

    Public Enum FolderReviewContext
        ForScanning = 1
        ForViewing = 2
    End Enum
    Friend gFolderReviewWindowContext As FolderReviewContext

    Friend PAFWSaysMWTopShouldBe As Double
    Friend PAFWSaysMWLeftShouldBe As Double

    Friend gDeletedFolderIndex As Integer = -1
    Friend gSentFolderIndex As Integer = -1
    Friend gInboxFolderIndex As Integer = -1

    Friend gContextFile1 As String = ""
    Friend gContextFile2 As String = ""
    Friend gContextFile3 As String = ""
    Friend gContextFile4 As String = ""

    Friend Enum WhoIsInControlType As Int16
        Main = 1
        PickAFolder = 2
    End Enum
    Friend gWhoIsInControl As WhoIsInControlType = WhoIsInControlType.PickAFolder

    Friend gPickFromContextMenuOverride As Integer = -1

    'Shared Routines

#Region "Custom Message Box"

    Friend Function ShowMessageBox(Optional ByVal Caption As String = "",
                              Optional ByVal InstructionalIcon As CustomDialog.CustomDialogIcons = CustomDialog.CustomDialogIcons.None,
                               Optional ByVal InstructionHeading As String = "",
                               Optional ByVal InstructionText As String = "",
                               Optional ByVal AdditionalDetails As String = "",
                               Optional ByVal FooterText As String = "",
                               Optional ByVal FooterIcon As CustomDialog.CustomDialogIcons = CustomDialog.CustomDialogIcons.None,
                               Optional ByVal Button As CustomDialog.CustomDialogButtons = CustomDialog.CustomDialogButtons.OK,
                               Optional ByVal DefaultButton As CustomDialog.CustomDialogResults = CustomDialog.CustomDialogButtons.OK,
                               Optional ByVal intButtonsDisabledDelay As Integer = 0) As CustomDialog.CustomDialogResults

        Dim ReturnCode As CustomDialog.CustomDialogResults

        If My.Settings.SoundAlert Then
            If (InstructionalIcon = CustomDialog.CustomDialogIcons.Warning) Or (InstructionalIcon = CustomDialog.CustomDialogIcons.Stop) Then
                Beep()
            End If

        End If

        Dim gCustomDialog = New CustomDialog(Caption,
                                                InstructionHeading,
                                                InstructionText,
                                                AdditionalDetails,
                                                FooterText,
                                                Button,
                                                DefaultButton,
                                                InstructionalIcon,
                                                FooterIcon,
                                                intButtonsDisabledDelay)

        ReturnCode = gCustomDialog.Show
        gCustomDialog = Nothing

        Return ReturnCode

    End Function

#End Region

    Friend Function LookupFolderNamesTableIndex(ByVal str As String) As Integer

        str = str.TrimStart("\")
        If str.Length = 0 Then
            Return -1
        Else
            Return Array.IndexOf(gFolderNamesTableTrimmed, str)
        End If

    End Function

#Region "Process Key Strokes"

    Friend ShiftOn As Boolean = False
    Friend CtrlOn As Boolean = False
    Friend AltOn As Boolean = False

    Friend MenuKeyStrokeOverRide As Boolean = False
    Friend gProxyAction As String

    Friend Sub ProcessKeyUp(ByVal e As System.Windows.Input.KeyEventArgs)

        If (e.Key.ToString = "RightShift") Or (e.Key.ToString = "LeftShift") Then
            ShiftOn = False

        ElseIf (e.Key.ToString = "RightCtrl") Or (e.Key.ToString = "LeftCtrl") Then
            CtrlOn = False

        ElseIf (e.Key.ToString = "RightAlt") Or (e.Key.ToString = "LeftAlt") Then
            AltOn = False

        End If

    End Sub

    Friend Sub ProcessKeyDown(ByVal e As System.Windows.Input.KeyEventArgs)

        If (e.Key = Key.Z) And CtrlOn Then
            gProxyAction = "Undo"
            gMainWindow.SafelyPerformActionByProxy()
            Exit Sub
        End If
        'Console.WriteLine(e.Key.ToString)

        If (e.Key = Key.RightAlt) Or (e.Key = Key.LeftAlt) Or (e.Key = 156) Then
            MenuKeyStrokeOverRide = Not MenuKeyStrokeOverRide
            If MenuKeyStrokeOverRide Then
                gMainWindow.SafelyActivateMenu()
            End If
            Exit Sub
        End If

        gSentText = ""

        If MenuKeyStrokeOverRide Then Exit Sub

        If e.Key = Key.F5 Then
            gBypassRefreshPrompt = True
            gProxyAction = "Refresh"
            gMainWindow.SafelyPerformActionByProxy()
            Exit Sub
        End If

        Select Case e.Key.ToString

            Case Is = "Delete"
                gProxyAction = "Delete"
                gMainWindow.SafelyPerformActionByProxy()

            Case Is = "Escape"
                gSentText = "Escape"

            Case "RightShift", "LeftShift"
                ShiftOn = True

            Case "RightCtrl", "LeftCtrl"
                CtrlOn = True

            Case "RightAlt", "LeftAlt"
                AltOn = True

            Case Is = "Back"
                gSentText = vbBack

            Case "D0" To "D9", "NumPad0" To "NumPad9"

                If e.Key.ToString.Length = 2 Then
                    gSentText = e.Key.ToString.Remove(0, 1)
                Else
                    gSentText = e.Key.ToString.Remove(0, 6)
                End If

                If ShiftOn Then
                    If gSentText = "8" Then ' star is not allowed in a folder name
                        gSentText = ""
                    Else
                        Dim ws() As String = {")", "!", "@", "#", "$", "%", "^", "&", "*", "("}
                        gSentText = ws(CType(gSentText, Integer))
                    End If
                End If

            Case "OemPeriod", "Decimal"
                gSentText = "."

            Case "Subtract"
                gSentText = "-"

            Case "OemMinus"
                If ShiftOn Then
                    gSentText = "_"
                Else
                    gSentText = "-"
                End If

            Case "Add"
                gSentText = "+"

            Case "OemPlus"
                If ShiftOn Then
                    gSentText = "+"
                Else
                    gSentText = "="
                End If

            Case "Oem4"
                If ShiftOn Then
                    gSentText = "{"
                Else
                    gSentText = "["
                End If

            Case "Oem6"
                If ShiftOn Then
                    gSentText = "}"
                Else
                    gSentText = "]"
                End If

            Case "OemTilde"
                If ShiftOn Then
                    gSentText = "~"
                Else
                    gSentText = "`"
                End If

            Case "A" To "Z"
                If ShiftOn Then
                    gSentText = e.Key.ToString
                Else
                    gSentText = e.Key.ToString.ToLower
                End If

        End Select

        If (gSentText.Length = 1) Or (gSentText = "Escape") Then
            gPickAFolderWindow.SafelyUpdateQuickFilter()
        End If

    End Sub

#End Region

#Region "Make A Form Top Most"

    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer

    Friend MainWindow As IntPtr
    Friend PickAWindowHandle As IntPtr
    Friend CustomDialogWindowHandle As IntPtr

    Friend Sub MakeTopMost(ByVal MakeTopMostFlag As Boolean, ByVal WindowHandle As IntPtr)

        'Me.BringIntoView()
        'Me.Focus()
        'Me.Activate()

        Dim HWND_TOPMOST As Integer
        If MakeTopMostFlag Then
            HWND_TOPMOST = -1
        Else
            HWND_TOPMOST = -2
        End If

        Dim SWP_NOMOVE As Integer = &H2
        Dim SWP_NOSIZE As Integer = &H1
        Dim TOPMOST_FLAGS As Integer = SWP_NOMOVE Or SWP_NOSIZE
        Dim hwnd As Integer = WindowHandle.ToInt64
        Try
            SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS)
        Catch ex As Exception
        End Try

    End Sub

#End Region

#Region "Check for Update"

    Friend Sub CheckIfNewVersionIsAvailable()

        If Not My.Settings.UpgradeNotify Then Exit Sub

        ' Deal with situations where an upgrade check should be forced
        Dim ForceAnUpgradeCheck As Boolean = False

        Try

            ' when the NextUpgradeCheckDate setting is not present
            ForceAnUpgradeCheck = (My.Settings.NextUpgradeCheckDate = Nothing)

            ' when NextUpgradeCheckDate is not a valid length
            If ForceAnUpgradeCheck Then
            Else
                ForceAnUpgradeCheck = My.Settings.NextUpgradeCheckDate.Length <> "yyyyMMdd".Length()
            End If

            ' when NextUpgradeCheckDate is not a valid date
            If ForceAnUpgradeCheck Then
            Else
                Dim TestDate As DateTime
                ForceAnUpgradeCheck = Not DateTime.TryParseExact(My.Settings.NextUpgradeCheckDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, TestDate)
            End If

            ' when NextUpgradeCheckDate is more than nine days in the future (by default it should never be more than 7 days in the future)
            If ForceAnUpgradeCheck Then
            Else
                Dim TestDate As DateTime = DateTime.ParseExact(My.Settings.NextUpgradeCheckDate, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
                ForceAnUpgradeCheck = TestDate > Today.AddDays(9)
            End If

        Catch ex As Exception

            ForceAnUpgradeCheck = True

        End Try


        ' check weekly

        If ForceAnUpgradeCheck OrElse ((Format(Today, "yyyyMMdd") >= My.Settings.NextUpgradeCheckDate)) Then

            Try


                Dim ThisVersion As ThisVersionIs = CheckFilefriendlyVersion()

                Select Case ThisVersion

                    Case Is = ThisVersionIs.CanNotTellAtThisTime
                    ' do nothing

                    Case Is = ThisVersionIs.Current
                    ' do nothing

                    Case Is = ThisVersionIs.ABeta
                    ' do nothing

                    Case Is = ThisVersionIs.OutOfDate

                        If ShowMessageBox("FileFriendly - Check for Update",
                            CustomDialog.CustomDialogIcons.Information,
                            "A newer version is available.",
                            "Would you like to go to the download page now?",
                            "If you would like to you can turn this pop-up window off in the Options Window.",
                            "",
                            CustomDialog.CustomDialogIcons.None,
                            CustomDialog.CustomDialogButtons.YesNo,
                            CustomDialog.CustomDialogResults.Yes) = CustomDialog.CustomDialogResults.Yes Then

                            System.Diagnostics.Process.Start(gDownloadWebPage)

                        End If

                End Select

                My.Settings.NextUpgradeCheckDate = Format(Today.AddDays(7), "yyyyMMdd")
                My.Settings.Save()

            Catch ex As Exception
            End Try

        End If

    End Sub

    Friend Enum ThisVersionIs
        OutOfDate = 0
        Current = 1
        ABeta = 2
        CanNotTellAtThisTime = 3
    End Enum
    Friend Function CheckFilefriendlyVersion() As ThisVersionIs

        'supports version format of up to: 99.99.99.99

        Dim ReturnCode As ThisVersionIs = ThisVersionIs.CanNotTellAtThisTime

        Try

            Dim strCurrentVersionRunning() As String = System.Windows.Forms.Application.ProductVersion.ToString.Split(".")
            Dim intCurrentVersionRunning(3) As Integer
            intCurrentVersionRunning(0) = CType(strCurrentVersionRunning(0), Integer)
            intCurrentVersionRunning(1) = CType(strCurrentVersionRunning(1), Integer)
            intCurrentVersionRunning(2) = CType(strCurrentVersionRunning(2), Integer)
            intCurrentVersionRunning(3) = CType(strCurrentVersionRunning(3), Integer)

            Dim strContentsOfWebFile As String = ""

            If strContentsOfWebFile.Length = 0 Then
                Dim myWebClient As New System.Net.WebClient
                Dim file As New System.IO.StreamReader(myWebClient.OpenRead(gCurrentVersionWebPage))

                strContentsOfWebFile = file.ReadToEnd()

                file.Close()
                file.Dispose()

                myWebClient.Dispose()

            End If

            If strContentsOfWebFile.Length = 0 Then Exit Try

            Dim AllEntries() As String = Split(strContentsOfWebFile, vbCrLf)

            Try

                'Current version should be the top most record
                Dim TopMostEntry As String = AllEntries(0)
                Dim LineItems() As String
                Dim strMostCurrentVersionOnFile() As String
                Dim intCurrentVersionOnFile(3) As Integer

                'Top most record should say 
                'version x.x.x.x 
                If TopMostEntry.StartsWith("version") Then
                    LineItems = TopMostEntry.Split(" ")
                    strMostCurrentVersionOnFile = LineItems(1).Split(".")
                    intCurrentVersionOnFile(0) = CType(strMostCurrentVersionOnFile(0), Integer)
                    intCurrentVersionOnFile(1) = CType(strMostCurrentVersionOnFile(1), Integer)
                    intCurrentVersionOnFile(2) = CType(strMostCurrentVersionOnFile(2), Integer)
                    intCurrentVersionOnFile(3) = CType(strMostCurrentVersionOnFile(3), Integer)
                Else
                    Exit Try
                End If

                Dim lCurrentVersionRunning As Long =
                intCurrentVersionRunning(0) * 1000000 +
                intCurrentVersionRunning(1) * 10000 +
                intCurrentVersionRunning(2) * 100 +
                intCurrentVersionRunning(3)

                Dim lCurrentVersionOnFile As Long =
                intCurrentVersionOnFile(0) * 1000000 +
                intCurrentVersionOnFile(1) * 10000 +
                intCurrentVersionOnFile(2) * 100 +
                intCurrentVersionOnFile(3)

                If lCurrentVersionOnFile = lCurrentVersionRunning Then
                    ReturnCode = ThisVersionIs.Current
                ElseIf lCurrentVersionOnFile > lCurrentVersionRunning Then
                    ReturnCode = ThisVersionIs.OutOfDate
                Else
                    ReturnCode = ThisVersionIs.ABeta
                End If

            Catch ex As Exception
            End Try

        Catch ex As Exception
        End Try

        Return ReturnCode

    End Function

    <System.Diagnostics.DebuggerStepThrough()> Friend Function QuickFilter(ByVal InputString As String, ByVal Filter As String) As String

        Dim OutputValue As String = ""
        Try
            Dim x As Int32
            For x = 1 To Len(InputString)
                If Filter.IndexOf(Mid(InputString, x, 1)) = -1 Then
                Else
                    OutputValue &= Mid(InputString, x, 1)
                End If
            Next
        Catch ex As Exception
        End Try
        Return OutputValue

    End Function

#End Region

#Region "Memory Management"

    Public Class MemoryManagement

        Private Declare Function SetProcessWorkingSetSize Lib "kernel32.dll" (
          ByVal process As IntPtr,
          ByVal minimumWorkingSetSize As Integer,
          ByVal maximumWorkingSetSize As Integer) As Integer

        Public Shared Sub FlushMemory()

            Try

                If (Environment.OSVersion.Platform = PlatformID.Win32NT) Then
                    Dim p As Process = Process.GetCurrentProcess
                    SetProcessWorkingSetSize(p.Handle, -1, -1)
                    p.Dispose()
                Else
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                End If

            Catch ex As Exception

            End Try
        End Sub

    End Class

#End Region



End Module
