Imports System.ComponentModel

Namespace TreeViewWithCheckBoxes

    Public Class FooViewModel
        Implements INotifyPropertyChanged

        ' the FooViewModel provides the logic and data structure for a tree view with checkboxes, enabling folder selection and state propagation in a WPF UI.

#Region "Data"

        Public _Parent As FooViewModel
        Public _Name As String
        Public _FullPathName As String
        Public _Children As List(Of FooViewModel)
        Public _IsInitiallySelected As System.Nullable(Of Boolean) = False
        Public _isChecked As System.Nullable(Of Boolean) = False

        Structure WorkingTableStructure
            Dim Level As Integer
            Dim FolderName As String
            Dim FullPathName As String
            Dim IsChecked As System.Nullable(Of Boolean)
        End Structure
        Private Shared WorkingTable(gFolderNamesTable.Length) As WorkingTableStructure

#End Region

#Region "CreateFoos"

        Public Shared Function CreateFoos() As List(Of FooViewModel)

            Dim ReturnValues As List(Of FooViewModel) = New List(Of FooViewModel)

            Try

                Dim strCollectionOfExcludedFolders As System.Collections.Specialized.StringCollection = New System.Collections.Specialized.StringCollection

                If gFolderReviewWindowContext = FolderReviewContext.ForViewing Then
                    strCollectionOfExcludedFolders = My.Settings.ExcludedViewFolders
                Else
                    strCollectionOfExcludedFolders = My.Settings.ExcludedScanFolders
                End If

                'Table Prep

                Dim WorkingFoldersNameTable(gFolderTable.Length - 1) As String
                Array.Copy(gFolderNamesTable, WorkingFoldersNameTable, gFolderNamesTable.Length)
                Array.Sort(WorkingFoldersNameTable)

                Dim CurrentFolder() As String
                Dim CurrentFolderName As String
                Dim CurrentLevel As Integer

                For i As Integer = 0 To gFolderNamesTable.Length - 1

                    CurrentFolder = WorkingFoldersNameTable(i).Split("\")
                    CurrentLevel = CurrentFolder.Length - 1
                    CurrentFolderName = CurrentFolder(CurrentLevel)
                    CurrentLevel -= 1

                    WorkingTable(i).Level = CurrentLevel
                    WorkingTable(i).FolderName = CurrentFolderName
                    WorkingTable(i).FullPathName = WorkingFoldersNameTable(i)

                    If strCollectionOfExcludedFolders IsNot Nothing Then
                        If strCollectionOfExcludedFolders.IndexOf(WorkingTable(i).FullPathName) > -1 Then
                            WorkingTable(i).IsChecked = False
                        Else
                            WorkingTable(i).IsChecked = True
                        End If
                    Else
                        WorkingTable(i).IsChecked = True
                    End If

                Next

                WorkingTable(gFolderNamesTable.Length).Level = -1
                WorkingTable(gFolderNamesTable.Length).FolderName = "*end*"
                WorkingTable(gFolderNamesTable.Length).FullPathName = "*end*"


                If gFolderReviewWindowContext = FolderReviewContext.ForScanning Then
                    If My.Settings.FirstRun Then

                        My.Settings.FirstRun = False
                        My.Settings.Save()

                        ' Make sure none of the excluded folders are checked on initial load
                        Dim ExcludeFolders As String = "\\Outlook\Inbox \\Outlook\Outbox \\Outlook\Deleted Items \\Outlook\Drafts \\Outlook\Sent Items \\Outlook\Spam \\Outlook\Junk E-mail \\Outlook\RSS Feeds "
                        For z = 0 To WorkingTable.Length - 1
                            If ExcludeFolders.Contains(WorkingTable(z).FullPathName) Then
                                WorkingTable(z).IsChecked = False
                            End If
                        Next

                        'Exclude Archived Folders
                        For z = 0 To WorkingTable.Length - 1
                            If WorkingTable(z).FullPathName.StartsWith("\\Archive") Then
                                WorkingTable(z).IsChecked = False
                            End If
                        Next

                    End If

                End If

                Dim RootShouldBeChecked As Boolean = True

                If strCollectionOfExcludedFolders Is Nothing Then
                    RootShouldBeChecked = True
                Else
                    If strCollectionOfExcludedFolders.Count = 0 Then
                        RootShouldBeChecked = True
                    ElseIf strCollectionOfExcludedFolders.Count = 1 Then
                        If strCollectionOfExcludedFolders(0).ToString = "*start*" Then
                            RootShouldBeChecked = True
                        Else
                            RootShouldBeChecked = False
                        End If
                    Else
                        RootShouldBeChecked = False
                    End If
                End If

                Dim MasterChildList As List(Of FooViewModel) = ReturnAllChildren(0, WorkingTable.Length - 1)
                Dim Root As New FooViewModel("All Folders", "*All Folders*", RootShouldBeChecked) With {.IsInitiallySelected = False, .Children = MasterChildList}
                Root.FullPathName = "*start*"

                Root.IsInitiallySelected = True

                ReturnValues.Add(Root)

            Catch ex As Exception
            End Try

            Return ReturnValues

        End Function


        Public Shared Function ReturnAllChildren(ByVal StartRecord As Integer, ByRef EndRecord As Integer) As List(Of FooViewModel)

            Static Dim MostAdvancedCounter As Integer

            If StartRecord = 0 Then MostAdvancedCounter = 0

            Dim ChildList As New List(Of FooViewModel)

            For i = StartRecord To EndRecord

                If MostAdvancedCounter > i Then i = MostAdvancedCounter

                If WorkingTable(i).Level = WorkingTable(StartRecord).Level Then

                    If WorkingTable(i + 1).Level > WorkingTable(StartRecord).Level Then
                        'use recursion to find additional child records under current record
                        ChildList.Add(New FooViewModel(WorkingTable(i).FolderName, WorkingTable(i).FullPathName, WorkingTable(i).IsChecked) With {.IsInitiallySelected = False, .Children = ReturnAllChildren(i + 1, EndRecord)})
                    Else
                        ChildList.Add(New FooViewModel(WorkingTable(i).FolderName, WorkingTable(i).FullPathName, WorkingTable(i).IsChecked))
                    End If

                Else

                    ' returning from recursions
                    MostAdvancedCounter = i
                    Exit For

                End If

            Next

            Return ChildList

        End Function

        Private Sub New(ByVal name As String, ByVal fullpathname As String, ByVal checked As System.Nullable(Of Boolean))

            Me.Name = name
            Me.FullPathName = fullpathname
            Me.IsChecked = checked
            Me.Children = New List(Of FooViewModel)

        End Sub

        Private Sub Initialize()

            For Each child As FooViewModel In Me.Children
                child._Parent = Me
                child.Initialize()
            Next child

            If Me.Children.Count > 0 Then
                VerifyCheckState()
            End If

        End Sub


#End Region


#Region "INotifyPropertyChanged Members"

        'original
        'Private Sub OnPropertyChanged(ByVal prop As String)
        '    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(prop))
        'End Sub
        'Public Event PropertyChanged As PropertyChangedEventHandler

        'Public Sub PropertyChanged(ByVal sender As Object, ByVal e As System.ComponentModel.INotifyPropertyChanged)
        '    'RaiseEvent xxPropertyChanged(Me, New PropertyChangedEventArgs(sender))
        'End Sub


        ' End Sub
        'Public Event PropertyChanged As PropertyChangedEventHandler


        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Public Sub NotifyPropertyChanged(ByVal prop As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(prop))
        End Sub

#End Region

#Region "Properties"

        Public Property Children() As List(Of FooViewModel)
            Get
                Return Me._Children
            End Get
            Private Set(ByVal value As List(Of FooViewModel))
                Me._Children = value
            End Set
        End Property

        Public Property IsInitiallySelected() As Boolean
            Get
                Return Me._IsInitiallySelected
            End Get
            Set(ByVal value As Boolean)
                Me._IsInitiallySelected = value
            End Set
        End Property

        Public Property Name() As String
            Get
                Return Me._Name
            End Get
            Set(ByVal value As String)
                Me._Name = value
            End Set
        End Property


        Public Property FullPathName() As String
            Get
                Return Me._FullPathName
            End Get
            Set(ByVal value As String)
                Me._FullPathName = value
            End Set
        End Property

#Region "IsChecked"


        ' Gets/sets the state of the associated UI toggle (ex. CheckBox).
        ' The return value is calculated based on the check state of all
        ' child FooViewModels.  Setting this property to true or false
        ' will set all children to the same check state, and setting it 
        ' to any value will cause the parent to verify its check state.

        Public Property IsChecked() As System.Nullable(Of Boolean)
            Get
                Return _isChecked
            End Get
            Set(ByVal value As System.Nullable(Of Boolean))
                Me.SetIsChecked(value, True, True)
            End Set
        End Property

        Public Sub SetIsChecked(ByVal value As System.Nullable(Of Boolean), ByVal updateChildren As Boolean, ByVal updateParent As Boolean)

            If value = _isChecked Then
                Return
            End If

            _isChecked = value

            If Me.Children IsNot Nothing Then
                If updateChildren AndAlso _isChecked.HasValue Then
                    Me.Children.ForEach(Function(c As Object) c.SetIsChecked(_isChecked, True, False))
                End If
            End If

            If updateParent AndAlso _Parent IsNot Nothing Then
                _Parent.VerifyCheckState()
            End If

            'Me.OnPropertyChanged("IsChecked")
            Me.NotifyPropertyChanged("IsChecked")
        End Sub

        'Private Shared Sub CheckChildren(ByVal _isChecked As System.Nullable(Of Boolean), ByVal o As Object)

        '    Try

        '        If o Is Nothing Then Exit Try

        '        For i As Integer = 0 To o.count - 1
        '            Console.WriteLine(o(i).name.ToString)
        '            If o(i).name = "Bills" Then
        '                Dim x As Integer = 123
        '            End If

        '            If o(i).children Is Nothing Then
        '                o(i).IsChecked = _isChecked
        '            Else
        '                If o(i).children.count > 0 Then
        '                    CheckChildren(_isChecked, o(i).Children)
        '                    o(i).IsChecked = _isChecked
        '                Else
        '                    o(i).IsChecked = _isChecked
        '                End If
        '            End If

        '        Next

        '    Catch ex As Exception
        '    End Try

        'End Sub



        Public Sub VerifyCheckState()
            Dim state As System.Nullable(Of Boolean) = Nothing
            For i As Integer = 0 To Me.Children.Count - 1
                Dim current As System.Nullable(Of Boolean) = Me.Children(i).IsChecked
                If i = 0 Then
                    state = current
                    'ElseIf state IsNot current Then
                ElseIf state <> current Then
                    state = Nothing
                    Exit For
                End If
            Next
            Me.SetIsChecked(state, False, True)
        End Sub

#End Region

#End Region

    End Class

End Namespace

