Public Class CustomDialogWindow
    Inherits System.Windows.Window

    Public Enum CustomDialogButtons
        OK
        OKCancel
        YesNo
        YesNoCancel
    End Enum

    Public Enum CustomDialogIcons
        None
        Information
        Question
        Shield
        [Stop]
        Warning
    End Enum

    Public Enum CustomDialogResults
        None
        OK
        Cancel
        Yes
        No
    End Enum

#Region " Private Declarations "

    Private _bolAeroGlassEnabled As Boolean = False
    Private _enumCustomDialogResult As CustomDialogResults = CustomDialogResults.None
    Private _intButtonsDisabledDelay As Integer
    Private _objButtonDelayTimer As System.Windows.Forms.Timer

#End Region

#Region " Constructors "

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub New(ByVal intButtonsDisabledDelay As Integer)
        InitializeComponent()

        If System.Environment.OSVersion.Version.Major < 6 Then
            Me.AllowsTransparency = True
            _bolAeroGlassEnabled = False

        Else
            _bolAeroGlassEnabled = True
        End If

        _intButtonsDisabledDelay = intButtonsDisabledDelay

    End Sub

#End Region

#Region " Public Properties "

    Public ReadOnly Property CustomDialogResult() As CustomDialogResults
        Get
            Return _enumCustomDialogResult
        End Get
    End Property

#End Region

    Private Sub MainWindow_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Me.MouseLeftButtonDown
        On Error Resume Next
        DragMove()
    End Sub

#Region " Methods "

    Protected Overrides Sub OnSourceInitialized(ByVal e As System.EventArgs)
        MyBase.OnSourceInitialized(e)

        If _bolAeroGlassEnabled = False Then
            'no aero glass
            Me.borderCustomDialog.Background = System.Windows.SystemColors.ActiveCaptionBrush
            Me.tbCaption.Foreground = System.Windows.SystemColors.ActiveCaptionTextBrush
            Me.borderCustomDialog.CornerRadius = New CornerRadius(10, 10, 0, 0)
            Me.borderCustomDialog.Padding = New Thickness(4, 0, 4, 4)
            Me.borderCustomDialog.BorderThickness = New Thickness(0, 0, 1, 1)
            Me.borderCustomDialog.BorderBrush = System.Windows.Media.Brushes.Black

        End If

    End Sub

    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnCancel.Click
        _enumCustomDialogResult = CustomDialogResults.Cancel
        Me.DialogResult = True
    End Sub

    Private Sub btnNo_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnNo.Click
        _enumCustomDialogResult = CustomDialogResults.No
        Me.DialogResult = True
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnOK.Click
        _enumCustomDialogResult = CustomDialogResults.OK
        Me.DialogResult = True
    End Sub

    Private Sub btnYes_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnYes.Click
        _enumCustomDialogResult = CustomDialogResults.Yes
        Me.DialogResult = True
    End Sub

    Private Sub CustomDialogWindow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing

        'this prevents ALT-F4 from closing the dialog box
        If Me.DialogResult.HasValue AndAlso Me.DialogResult.Value = True Then
            e.Cancel = False
        Else
            e.Cancel = True
        End If

    End Sub

    Private Sub CustomDialogWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        Me.tbAdditionalDetailsText.Visibility = Windows.Visibility.Collapsed
        Me.tbAdditionalDetailsText.FontFamily = Me.tbInstructionText.FontFamily
        Me.tbAdditionalDetailsText.FontSize = Me.tbInstructionText.FontSize
        Me.tbAdditionalDetailsText.FontStyle = Me.tbInstructionText.FontStyle
        Me.tbAdditionalDetailsText.FontWeight = Me.tbInstructionText.FontWeight

        If Me.ResizeMode <> Windows.ResizeMode.NoResize Then
            'this work around is necessary when glass is enabled and the window style is None which removes the chrome because the resize mode MUST be set to CanResize or else glass won't display
            Me.MinHeight = Me.ActualHeight
            Me.MaxHeight = Me.ActualHeight

            Me.MinWidth = Me.ActualWidth
            Me.MaxWidth = Me.ActualWidth
        End If

        If _intButtonsDisabledDelay > 0 Then
            Me.pbDisabledButtonsProgressBar.Maximum = _intButtonsDisabledDelay
            Me.pbDisabledButtonsProgressBar.IsIndeterminate = False

            Dim objDuration As New Duration(TimeSpan.FromSeconds(_intButtonsDisabledDelay))
            Dim objDoubleAnimation As New System.Windows.Media.Animation.DoubleAnimation(_intButtonsDisabledDelay, objDuration)
            Me.pbDisabledButtonsProgressBar.BeginAnimation(ProgressBar.ValueProperty, objDoubleAnimation)
            btnCancel.IsEnabled = False
            btnNo.IsEnabled = False
            btnOK.IsEnabled = False
            btnYes.IsEnabled = False
            _objButtonDelayTimer = New System.Windows.Forms.Timer
            AddHandler _objButtonDelayTimer.Tick, AddressOf OnTimedEvent
            _objButtonDelayTimer.Interval = _intButtonsDisabledDelay * 1000
            _objButtonDelayTimer.Start()

        Else
            Me.pbDisabledButtonsProgressBar.Visibility = Windows.Visibility.Collapsed
        End If

        Dim hwndSource As System.Windows.Interop.HwndSource = TryCast(PresentationSource.FromVisual(Me), System.Windows.Interop.HwndSource)
        If hwndSource IsNot Nothing Then
            CustomDialogWindowHandle = hwndSource.Handle
            MakeTopMost(True, CustomDialogWindowHandle)
        End If

    End Sub

    Private Sub expAdditionalDetails_Collapsed(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles expAdditionalDetails.Collapsed
        Me.expAdditionalDetails.Header = "See Details"
        Me.tbAdditionalDetailsText.Visibility = Windows.Visibility.Collapsed
        Me.btnCopyDetails.Visibility = Windows.Visibility.Collapsed
        Me.UpdateLayout()

        If Me.ResizeMode <> Windows.ResizeMode.NoResize Then
            Me.MaxHeight = Me.ActualHeight
        End If

    End Sub

    Private Sub expAdditionalDetails_Expanded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles expAdditionalDetails.Expanded

        If Me.ResizeMode <> Windows.ResizeMode.NoResize Then
            Me.MaxHeight = Double.PositiveInfinity
        End If

        Me.expAdditionalDetails.Header = "Hide Details"
        Me.tbAdditionalDetailsText.Visibility = Windows.Visibility.Visible
        If String.IsNullOrEmpty(Me.tbAdditionalDetailsText.Text) = False Then
            Me.btnCopyDetails.Visibility = Windows.Visibility.Visible
        End If
        Me.UpdateLayout()

        If Me.ResizeMode <> Windows.ResizeMode.NoResize Then
            Me.MaxHeight = Me.ActualHeight
        End If

    End Sub

    Private Sub btnCopyDetails_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles btnCopyDetails.Click
        If String.IsNullOrEmpty(Me.tbAdditionalDetailsText.Text) = False Then

            Dim textToCopy As String = Me.tbCaption.Text & Me.tbInstructionText.Text & vbCrLf & vbCrLf & Me.tbAdditionalDetailsText.Text

            System.Windows.Clipboard.SetText(textToCopy)

        End If
    End Sub

    Private Sub OnTimedEvent(ByVal source As Object, ByVal e As EventArgs)
        _objButtonDelayTimer.Stop()
        _objButtonDelayTimer.Dispose()
        _objButtonDelayTimer = Nothing
        btnCancel.IsEnabled = True
        btnNo.IsEnabled = True
        btnOK.IsEnabled = True
        btnYes.IsEnabled = True
        Me.pbDisabledButtonsProgressBar.Visibility = Windows.Visibility.Collapsed
    End Sub

    Private Sub tbCaption_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles tbCaption.MouseLeftButtonDown
        DragMove()
    End Sub

#End Region

    Private Sub CustomDialogWindow_SizeChanged(ByVal sender As Object, ByVal e As System.Windows.SizeChangedEventArgs) Handles Me.SizeChanged
        Me.UpdateLayout()
    End Sub

    Private Sub CustomDialogWindow_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Input.KeyEventArgs) Handles Me.PreviewKeyDown
        If e.Key <> System.Windows.Input.Key.Left AndAlso e.Key <> System.Windows.Input.Key.Right Then Return

        Dim buttons As New List(Of System.Windows.Controls.Button)

        If btnYes.IsVisible AndAlso btnYes.IsEnabled Then buttons.Add(btnYes)
        If btnNo.IsVisible AndAlso btnNo.IsEnabled Then buttons.Add(btnNo)
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
