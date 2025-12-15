
''' <summary>
'''     Displays a modal custom dialog box and returns the result to the caller.  Logs the dialog box contents and users reply.
''' </summary>
Public Class CustomDialog

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

    Private _enumButtons As CustomDialogWindow.CustomDialogButtons = CustomDialogWindow.CustomDialogButtons.OK
    Private _enumCustomDialogResult As CustomDialogWindow.CustomDialogResults = CustomDialogWindow.CustomDialogResults.None
    Private _enumDefaultButton As CustomDialogWindow.CustomDialogResults = CustomDialogWindow.CustomDialogResults.None
    Private _enumFooterIcon As CustomDialogWindow.CustomDialogIcons = CustomDialogWindow.CustomDialogIcons.None
    Private _enumInstructionIcon As CustomDialogWindow.CustomDialogIcons = CustomDialogWindow.CustomDialogIcons.None

    Private _intButtonsDisabledDelay As Integer = 0
    Private _strAdditionalDetailsText As String = String.Empty
    Private _strCallingMethodName As String = String.Empty          'example:  "btnDashboard_Click"
    Private _strCallingModuleName As String = String.Empty          'example:  "CustomDialogSample.exe"
    Private _strCallingReflectedTypeName As String = String.Empty   'example:  "ApplicationMainWindow"	
    Private _strCaption As String = String.Empty
    Private _strFooterText As String = String.Empty
    Private _strInstructionHeading As String = String.Empty
    Private _strInstructionText As String = String.Empty

#End Region

#Region " Public Properties "

    ''' <summary>
    ''' Gets or sets the optional text is displayed to the user when they click to Show Details expander.  Used to provide a detailed explaination to the user.
    ''' </summary>
    Public Property AdditionalDetailsText() As String
        Get
            Return _strAdditionalDetailsText
        End Get
        Set(ByVal Value As String)
            _strAdditionalDetailsText = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the buttons that will be displayed.  The default is the OK button.
    ''' </summary>
    Public Property Buttons() As CustomDialogWindow.CustomDialogButtons
        Get
            Return _enumButtons
        End Get
        Set(ByVal Value As CustomDialogWindow.CustomDialogButtons)
            _enumButtons = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the number of seconds that the buttons will be disabled, providing time for the user to read the dialog before dismissing it.  Assigning a value also causes a progress bar to display the elapsed time before the buttons are enabled.
    ''' </summary>
    Public Property ButtonsDisabledDelay() As Integer
        Get
            Return _intButtonsDisabledDelay
        End Get
        Set(ByVal Value As Integer)
            _intButtonsDisabledDelay = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the calling method name.  If not set, the value will be determined from the stack frame and the calling method name will be used.  Normally this value is not set by the calling code.
    ''' </summary>
    Public Property CallingMethodName() As String
        Get
            Return _strCallingMethodName
        End Get
        Set(ByVal Value As String)
            _strCallingMethodName = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the calling module name.  If not set, the value will be determined from the stack frame and the calling module name will be used.  Normally this value is not set by the calling code.
    ''' </summary>
    Public Property CallingModuleName() As String
        Get
            Return _strCallingModuleName
        End Get
        Set(ByVal Value As String)
            _strCallingModuleName = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the calling type name.  If not set, the value will be determined from the stack frame and the calling reflected type name will be used.  Normally this value is not set by the calling code.
    ''' </summary>
    Public Property CallingReflectedTypeName() As String
        Get
            Return _strCallingReflectedTypeName
        End Get
        Set(ByVal Value As String)
            _strCallingReflectedTypeName = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the dialog box window caption.  The caption is displayed in the window chrome.
    ''' </summary>
    Public Property Caption() As String
        Get
            Return _strCaption
        End Get
        Set(ByVal Value As String)
            _strCaption = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the default button for the dialog box.  This value defaults to none.
    ''' </summary>
    Public Property DefaultButton() As CustomDialogWindow.CustomDialogResults
        Get
            Return _enumDefaultButton
        End Get
        Set(ByVal Value As CustomDialogWindow.CustomDialogResults)
            _enumDefaultButton = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the icon displayed in the dialog footer.  This values defaults to none.
    ''' </summary>
    Public Property FooterIcon() As CustomDialogWindow.CustomDialogIcons
        Get
            Return _enumFooterIcon
        End Get
        Set(ByVal Value As CustomDialogWindow.CustomDialogIcons)
            _enumFooterIcon = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the optional text displayed in the footer.
    ''' </summary>
    Public Property FooterText() As String
        Get
            Return _strFooterText
        End Get
        Set(ByVal Value As String)
            _strFooterText = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the heading text displayed in the dialog box.
    ''' </summary>
    Public Property InstructionHeading() As String
        Get
            Return _strInstructionHeading
        End Get
        Set(ByVal Value As String)
            _strInstructionHeading = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the icon displayed to the left of the instruction text.  This value defaults to none.
    ''' </summary>
    Public Property InstructionIcon() As CustomDialogWindow.CustomDialogIcons
        Get
            Return _enumInstructionIcon
        End Get
        Set(ByVal Value As CustomDialogWindow.CustomDialogIcons)
            _enumInstructionIcon = Value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the text message for the dialog.
    ''' </summary>
    Public Property InstructionText() As String
        Get
            Return _strInstructionText
        End Get
        Set(ByVal Value As String)
            _strInstructionText = Value
        End Set
    End Property

#End Region

#Region " Constructors "


    Public Sub New(ByVal strCaption As String, ByVal strInstructionHeading As String, ByVal strInstructionText As String, ByVal strAdditionalDetailsText As String, ByVal strFooterText As String, ByVal enumButtons As CustomDialogWindow.CustomDialogButtons, ByVal enumDefaultButton As CustomDialogWindow.CustomDialogResults, ByVal enumInstructionIcon As CustomDialogWindow.CustomDialogIcons, ByVal enumFooterIcon As CustomDialogWindow.CustomDialogIcons, Optional ByVal intButtonsDisabledDelay As Integer = 0)

        _strCaption = strCaption
        _strInstructionHeading = strInstructionHeading
        _strInstructionText = strInstructionText
        _strAdditionalDetailsText = strAdditionalDetailsText
        _strFooterText = strFooterText
        _enumButtons = enumButtons
        _enumDefaultButton = enumDefaultButton
        _enumInstructionIcon = enumInstructionIcon
        _enumFooterIcon = enumFooterIcon
        _intButtonsDisabledDelay = intButtonsDisabledDelay

    End Sub

#End Region

#Region " Methods "

    ''' <summary>
    '''     Shows the custom dialog described by the constructor and properties set by the caller, returns CustomDialogResults.
    ''' </summary>
    ''' <returns>
    '''     A emGovPower.Core.CommonDialog.CustomDialogResults value.
    ''' </returns>
    Public Function Show() As CustomDialogWindow.CustomDialogResults

        'get the calling code information
        Dim objTrace As System.Diagnostics.StackTrace = New System.Diagnostics.StackTrace()
        If _strCallingReflectedTypeName.Length = 0 Then
            _strCallingReflectedTypeName = objTrace.GetFrame(1).GetMethod.ReflectedType.Name
        End If

        If _strCallingMethodName.Length = 0 Then
            _strCallingMethodName = objTrace.GetFrame(1).GetMethod.Name
        End If

        If _strCallingModuleName.Length = 0 Then
            _strCallingModuleName = objTrace.GetFrame(1).GetMethod.Module.Name
        End If

        Dim obj As New CustomDialogWindow(Me.ButtonsDisabledDelay)

        obj.tbCaption.Text = Me.Caption

        If Me.FooterText.Length > 0 Then
            obj.tbFooterText.Text = Me.FooterText

            If Me.FooterIcon <> CustomDialogWindow.CustomDialogIcons.None Then
                obj.imgFooterIcon.Source = New BitmapImage(GetIcon(Me.FooterIcon))
            Else
                obj.imgFooterIcon.Visibility = Visibility.Collapsed
            End If

        Else
            obj.tbFooterText.Visibility = Visibility.Collapsed
            obj.imgFooterIcon.Visibility = Visibility.Collapsed
        End If

        If Me.InstructionIcon = CustomDialogWindow.CustomDialogIcons.None Then
            obj.imgInstructionIcon.Visibility = Visibility.Collapsed
        Else
            obj.imgInstructionIcon.Source = New BitmapImage(GetIcon(Me.InstructionIcon))
        End If

        obj.tbInstructionText.Text = Me.InstructionText

        obj.tbInstructionHeading.Text = Me.InstructionHeading

        If Me.AdditionalDetailsText.Length > 0 Then
            obj.tbAdditionalDetailsText.Text = Me.AdditionalDetailsText
        Else
            obj.expAdditionalDetails.Visibility = Visibility.Collapsed
        End If

        SetButtons(obj)

        obj.ShowInTaskbar = False

        obj.ShowDialog()
        _enumCustomDialogResult = obj.CustomDialogResult

        LogDialog()

        Return _enumCustomDialogResult
    End Function

    Private Function GetIcon(ByVal enumCustomDialogIcon As CustomDialogWindow.CustomDialogIcons) As Uri

        Select Case enumCustomDialogIcon

            Case CustomDialogWindow.CustomDialogIcons.Information
                Return New Uri("Resources/CustomDialogInformation.png", UriKind.Relative)

            Case CustomDialogWindow.CustomDialogIcons.None
                Return Nothing

            Case CustomDialogWindow.CustomDialogIcons.Question
                Return New Uri("Resources/CustomDialogQuestion.png", UriKind.Relative)

            Case CustomDialogWindow.CustomDialogIcons.Shield
                Return New Uri("Resources/CustomDialogShield.png", UriKind.Relative)

            Case CustomDialogWindow.CustomDialogIcons.Stop
                Return New Uri("Resources/CustomDialogStop.png", UriKind.Relative)

            Case CustomDialogWindow.CustomDialogIcons.Warning
                Return New Uri("Resources/CustomDialogWarning.png", UriKind.Relative)

            Case Else
                Throw New ArgumentOutOfRangeException("enumCustomDialogIcon", enumCustomDialogIcon.ToString, "Programmer did not program for this icon.")
        End Select

        Return Nothing
    End Function

    Private Sub LogDialog()

        'TODO - developers, you can log the result here.  There is rich information within this class to provides great tracking of your users dialog box activities.
        'ensure that you review each property and include them in your log entry
        'don't forget to log the Windows user name also

    End Sub

    Private Sub SetButtons(ByVal obj As CustomDialogWindow)

        Select Case Me.Buttons

            Case CustomDialogWindow.CustomDialogButtons.OK
                obj.btnCancel.Visibility = Visibility.Collapsed
                obj.btnNo.Visibility = Visibility.Collapsed
                obj.btnYes.Visibility = Visibility.Collapsed

            Case CustomDialogWindow.CustomDialogButtons.OKCancel
                obj.btnNo.Visibility = Visibility.Collapsed
                obj.btnYes.Visibility = Visibility.Collapsed

            Case CustomDialogWindow.CustomDialogButtons.YesNo
                obj.btnOK.Visibility = Visibility.Collapsed
                obj.btnCancel.Visibility = Visibility.Collapsed

            Case CustomDialogWindow.CustomDialogButtons.YesNoCancel
                obj.btnOK.Visibility = Visibility.Collapsed

            Case Else
                Throw New ArgumentOutOfRangeException("Buttons", Me.Buttons.ToString, "Programmer did not program for this selection.")

        End Select

        Select Case Me.DefaultButton

            Case CustomDialogWindow.CustomDialogResults.Cancel
                obj.btnCancel.IsDefault = True

            Case CustomDialogWindow.CustomDialogResults.No
                obj.btnNo.IsDefault = True
                obj.btnCancel.IsCancel = True

            Case CustomDialogWindow.CustomDialogResults.None
                'do nothing
                obj.btnCancel.IsCancel = True

            Case CustomDialogWindow.CustomDialogResults.OK
                obj.btnOK.IsDefault = True
                obj.btnCancel.IsCancel = True

            Case CustomDialogWindow.CustomDialogResults.Yes
                obj.btnYes.IsDefault = True
                obj.btnCancel.IsCancel = True

            Case Else
                Throw New ArgumentOutOfRangeException("DefaultButton", Me.DefaultButton.ToString, "Programmer did not program for this selection.")

        End Select

    End Sub

#End Region

End Class

