Imports System.ComponentModel
Imports System.Windows.Controls.Primitives

Namespace ListViewLayout

    ' Attached behavior that manages proportional/fixed/fill columns.
    Public Class ListViewLayoutManager

        Public Shared ReadOnly EnabledProperty As DependencyProperty =
            DependencyProperty.RegisterAttached(
                "Enabled",
                GetType(Boolean),
                GetType(ListViewLayoutManager),
                New FrameworkPropertyMetadata(New PropertyChangedCallback(AddressOf OnLayoutManagerEnabledChanged)))

        Private ReadOnly _listView As ListView
        Private _scrollViewer As ScrollViewer
        Private _loaded As Boolean
        Private _resizing As Boolean
        Private _resizeCursor As Cursor
        Private _verticalScrollBarVisibility As ScrollBarVisibility = ScrollBarVisibility.Auto
        Private _autoSizedColumn As GridViewColumn

        Public Sub New(listView As ListView)
            If listView Is Nothing Then Throw New ArgumentNullException("listView")

            _listView = listView
            AddHandler _listView.Loaded, AddressOf ListViewLoaded
            AddHandler _listView.Unloaded, AddressOf ListViewUnloaded
        End Sub

        Public Shared Sub SetEnabled(d As DependencyObject, value As Boolean)
            d.SetValue(EnabledProperty, value)
        End Sub

        Public Shared Function GetEnabled(d As DependencyObject) As Boolean
            Return CBool(d.GetValue(EnabledProperty))
        End Function

        Private Shared Sub OnLayoutManagerEnabledChanged(d As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim lv = TryCast(d, ListView)
            If lv IsNot Nothing AndAlso CBool(e.NewValue) Then
                ' Just construct – manager hooks into events and keeps itself alive through delegates
                Dim mgr As New ListViewLayoutManager(lv)
            End If
        End Sub

        ' --- Visual tree wiring ------------------------------------------------

        Private Sub ListViewLoaded(sender As Object, e As RoutedEventArgs)
            RegisterEvents(_listView)
            InitColumns()
            DoResizeColumns()
            _loaded = True
        End Sub

        Private Sub ListViewUnloaded(sender As Object, e As RoutedEventArgs)
            If Not _loaded Then Return
            UnregisterEvents(_listView)
            _loaded = False
        End Sub

        Private Sub RegisterEvents(start As DependencyObject)
            Dim count As Integer = VisualTreeHelper.GetChildrenCount(start)
            For i As Integer = 0 To count - 1
                Dim child As Visual = TryCast(VisualTreeHelper.GetChild(start, i), Visual)

                If TypeOf child Is Thumb Then
                    Dim thumb = DirectCast(child, Thumb)
                    Dim column = FindParentColumn(thumb)
                    If column IsNot Nothing Then
                        AddHandler thumb.PreviewMouseMove, AddressOf ThumbPreviewMouseMove
                        AddHandler thumb.PreviewMouseLeftButtonDown, AddressOf ThumbPreviewMouseLeftButtonDown
                        DependencyPropertyDescriptor.FromProperty(GridViewColumn.WidthProperty, GetType(GridViewColumn)).
                            AddValueChanged(column, AddressOf GridColumnWidthChanged)
                    End If

                ElseIf TypeOf child Is GridViewColumnHeader Then
                    Dim header = DirectCast(child, GridViewColumnHeader)
                    AddHandler header.SizeChanged, AddressOf GridColumnHeaderSizeChanged

                ElseIf _scrollViewer Is Nothing AndAlso TypeOf child Is ScrollViewer Then
                    _scrollViewer = DirectCast(child, ScrollViewer)
                    AddHandler _scrollViewer.ScrollChanged, AddressOf ScrollViewerScrollChanged
                    _scrollViewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden
                    _scrollViewer.VerticalScrollBarVisibility = _verticalScrollBarVisibility
                End If

                RegisterEvents(child) ' recursive
            Next
        End Sub

        Private Sub UnregisterEvents(start As DependencyObject)
            Dim count As Integer = VisualTreeHelper.GetChildrenCount(start)
            For i As Integer = 0 To count - 1
                Dim child As Visual = TryCast(VisualTreeHelper.GetChild(start, i), Visual)

                If TypeOf child Is Thumb Then
                    Dim thumb = DirectCast(child, Thumb)
                    Dim column = FindParentColumn(thumb)
                    If column IsNot Nothing Then
                        RemoveHandler thumb.PreviewMouseMove, AddressOf ThumbPreviewMouseMove
                        RemoveHandler thumb.PreviewMouseLeftButtonDown, AddressOf ThumbPreviewMouseLeftButtonDown
                        DependencyPropertyDescriptor.FromProperty(GridViewColumn.WidthProperty, GetType(GridViewColumn)).
                            RemoveValueChanged(column, AddressOf GridColumnWidthChanged)
                    End If

                ElseIf TypeOf child Is GridViewColumnHeader Then
                    Dim header = DirectCast(child, GridViewColumnHeader)
                    RemoveHandler header.SizeChanged, AddressOf GridColumnHeaderSizeChanged

                ElseIf _scrollViewer Is Nothing AndAlso TypeOf child Is ScrollViewer Then
                    _scrollViewer = DirectCast(child, ScrollViewer)
                    RemoveHandler _scrollViewer.ScrollChanged, AddressOf ScrollViewerScrollChanged
                End If

                UnregisterEvents(child) ' recursive
            Next
        End Sub

        Private Function FindParentColumn(element As DependencyObject) As GridViewColumn
            While element IsNot Nothing
                Dim header = TryCast(element, GridViewColumnHeader)
                If header IsNot Nothing Then
                    Return header.Column
                End If
                element = VisualTreeHelper.GetParent(element)
            End While
            Return Nothing
        End Function

        Private Function FindColumnHeader(start As DependencyObject, column As GridViewColumn) As GridViewColumnHeader
            Dim count As Integer = VisualTreeHelper.GetChildrenCount(start)
            For i As Integer = 0 To count - 1
                Dim child As Visual = TryCast(VisualTreeHelper.GetChild(start, i), Visual)

                Dim header = TryCast(child, GridViewColumnHeader)
                If header IsNot Nothing AndAlso header.Column Is column Then
                    Return header
                End If

                Dim nested = FindColumnHeader(child, column)
                If nested IsNot Nothing Then
                    Return nested
                End If
            Next
            Return Nothing
        End Function

        ' --- Column initialization / resize -----------------------------------

        Private Sub InitColumns()
            Dim view = TryCast(_listView.View, GridView)
            If view Is Nothing Then Return

            For Each col In view.Columns
                If RangeColumn.IsRangeColumn(col) Then
                    Dim minWidth = RangeColumn.GetRangeMinWidth(col)
                    Dim maxWidth = RangeColumn.GetRangeMaxWidth(col)

                    If minWidth.HasValue OrElse maxWidth.HasValue Then
                        Dim header = FindColumnHeader(_listView, col)
                        If header Is Nothing Then Continue For

                        If minWidth.HasValue Then header.MinWidth = minWidth.Value
                        If maxWidth.HasValue Then header.MaxWidth = maxWidth.Value
                    End If
                End If
            Next
        End Sub

        Private Sub DoResizeColumns()
            If _resizing Then Return

            _resizing = True
            Try
                ResizeColumns()
            Finally
                _resizing = False
            End Try
        End Sub

        Private Sub ResizeColumns()
            Dim view = TryCast(_listView.View, GridView)
            If view Is Nothing OrElse view.Columns.Count = 0 Then Return

            Dim actualWidth As Double = Double.PositiveInfinity
            If _scrollViewer IsNot Nothing Then
                actualWidth = _scrollViewer.ViewportWidth
            End If
            If Double.IsInfinity(actualWidth) Then
                actualWidth = _listView.ActualWidth
            End If
            If Double.IsInfinity(actualWidth) OrElse actualWidth <= 0 Then Return

            Dim resizableRegionCount As Double = 0
            Dim otherColumnsWidth As Double = 0

            For Each col In view.Columns
                If ProportionalColumn.IsProportionalColumn(col) Then
                    resizableRegionCount += ProportionalColumn.GetProportionalWidth(col).Value
                Else
                    otherColumnsWidth += col.ActualWidth
                End If
            Next

            If resizableRegionCount <= 0 Then
                ' No proportional columns: optional "fill" behaviour; otherwise allow horizontal scroll
                _scrollViewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto

                ' Look for first fill column (if you later mark one with RangeColumn.IsFillColumn)
                Dim fillColumn As GridViewColumn = Nothing
                For Each col In view.Columns
                    If IsFillColumn(col) Then
                        fillColumn = col
                        Exit For
                    End If
                Next

                If fillColumn IsNot Nothing Then
                    Dim otherWithoutFill = otherColumnsWidth - fillColumn.ActualWidth
                    Dim fillWidth = actualWidth - otherWithoutFill
                    If fillWidth > 0 Then
                        Dim minWidth = RangeColumn.GetRangeMinWidth(fillColumn)
                        Dim maxWidth = RangeColumn.GetRangeMaxWidth(fillColumn)

                        Dim setWidth As Boolean = True
                        If minWidth.HasValue AndAlso fillWidth < minWidth.Value Then setWidth = False
                        If maxWidth.HasValue AndAlso fillWidth > maxWidth.Value Then setWidth = False

                        If setWidth Then
                            _scrollViewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden
                            fillColumn.Width = fillWidth
                        End If
                    End If
                End If

                Return
            End If

            Dim resizableColumnsWidth = actualWidth - otherColumnsWidth
            If resizableColumnsWidth <= 0 Then Return

            Dim regionWidth = resizableColumnsWidth / resizableRegionCount

            For Each col In view.Columns
                If ProportionalColumn.IsProportionalColumn(col) Then
                    col.Width = ProportionalColumn.GetProportionalWidth(col).Value * regionWidth
                End If
            Next
        End Sub

        Private Function IsFillColumn(col As GridViewColumn) As Boolean
            If col Is Nothing Then Return False
            Dim view = TryCast(_listView.View, GridView)
            If view Is Nothing OrElse view.Columns.Count = 0 Then Return False

            Dim isFill = RangeColumn.GetRangeIsFillColumn(col)
            Return isFill.HasValue AndAlso isFill.Value
        End Function

        Private Function SetRangeColumnToBounds(col As GridViewColumn) As Double
            Dim startWidth = col.Width

            Dim minWidth = RangeColumn.GetRangeMinWidth(col)
            Dim maxWidth = RangeColumn.GetRangeMaxWidth(col)

            If minWidth.HasValue AndAlso maxWidth.HasValue AndAlso minWidth > maxWidth Then
                Return 0
            End If

            If minWidth.HasValue AndAlso col.Width < minWidth.Value Then
                col.Width = minWidth.Value
            ElseIf maxWidth.HasValue AndAlso col.Width > maxWidth.Value Then
                col.Width = maxWidth.Value
            End If

            Return col.Width - startWidth
        End Function

        ' --- Event handlers ----------------------------------------------------

        Private Sub ThumbPreviewMouseMove(sender As Object, e As MouseEventArgs)
            Dim thumb = TryCast(sender, Thumb)
            If thumb Is Nothing Then Return

            Dim col = FindParentColumn(thumb)
            If col Is Nothing Then Return

            ' Suppress resizing for proportional, fixed and fill columns
            If ProportionalColumn.IsProportionalColumn(col) OrElse
               FixedColumn.IsFixedColumn(col) OrElse
               IsFillColumn(col) Then
                thumb.Cursor = Nothing
                Return
            End If

            If thumb.IsMouseCaptured AndAlso RangeColumn.IsRangeColumn(col) Then
                Dim minWidth = RangeColumn.GetRangeMinWidth(col)
                Dim maxWidth = RangeColumn.GetRangeMaxWidth(col)

                If minWidth.HasValue AndAlso maxWidth.HasValue AndAlso minWidth > maxWidth Then
                    Return
                End If

                If _resizeCursor Is Nothing Then
                    _resizeCursor = thumb.Cursor
                End If

                If minWidth.HasValue AndAlso col.Width <= minWidth.Value Then
                    thumb.Cursor = Cursors.[No]
                ElseIf maxWidth.HasValue AndAlso col.Width >= maxWidth.Value Then
                    thumb.Cursor = Cursors.[No]
                Else
                    thumb.Cursor = _resizeCursor
                End If
            End If
        End Sub

        Private Sub ThumbPreviewMouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
            Dim thumb = TryCast(sender, Thumb)
            If thumb Is Nothing Then Return

            Dim col = FindParentColumn(thumb)

            If ProportionalColumn.IsProportionalColumn(col) OrElse
               FixedColumn.IsFixedColumn(col) OrElse
               IsFillColumn(col) Then
                e.Handled = True
            End If
        End Sub

        Private Sub GridColumnWidthChanged(sender As Object, e As EventArgs)
            If Not _loaded Then Return

            Dim col = TryCast(sender, GridViewColumn)
            If col Is Nothing Then Return

            ' Ignore proportional/fixed columns
            If ProportionalColumn.IsProportionalColumn(col) OrElse FixedColumn.IsFixedColumn(col) Then
                Return
            End If

            If RangeColumn.IsRangeColumn(col) Then
                If Double.IsNaN(col.Width) Then
                    _autoSizedColumn = col
                    Return
                End If

                If SetRangeColumnToBounds(col) <> 0 Then
                    Return
                End If
            End If

            DoResizeColumns()
        End Sub

        Private Sub GridColumnHeaderSizeChanged(sender As Object, e As SizeChangedEventArgs)
            If _autoSizedColumn Is Nothing Then Return

            Dim header = TryCast(sender, GridViewColumnHeader)
            If header Is Nothing Then Return

            If header.Column Is _autoSizedColumn Then
                If Double.IsNaN(header.Width) Then
                    header.Column.Width = header.ActualWidth
                    DoResizeColumns()
                End If

                _autoSizedColumn = Nothing
            End If
        End Sub

        Private Sub ScrollViewerScrollChanged(sender As Object, e As ScrollChangedEventArgs)
            If _loaded AndAlso e.ViewportWidthChange <> 0 Then
                DoResizeColumns()
            End If
        End Sub

    End Class

    ' --- Base helper for attached props --------------------------------------

    Public MustInherit Class LayoutColumn
        Protected Shared Function HasPropertyValue(column As GridViewColumn, dp As DependencyProperty) As Boolean
            If column Is Nothing Then Throw New ArgumentNullException("column")
            Dim value = column.ReadLocalValue(dp)
            Return (value IsNot Nothing AndAlso value.[GetType]() Is dp.PropertyType)
        End Function

        Protected Shared Function GetColumnWidth(column As GridViewColumn, dp As DependencyProperty) As Double?
            If column Is Nothing Then Throw New ArgumentNullException("column")
            Dim value = column.ReadLocalValue(dp)
            If value IsNot Nothing AndAlso value.[GetType]() Is dp.PropertyType Then
                Return CDbl(value)
            End If
            Return Nothing
        End Function
    End Class

    ' --- Fixed columns (no proportional behaviour) ----------------------------

    Public NotInheritable Class FixedColumn
        Inherits LayoutColumn

        Public Shared ReadOnly WidthProperty As DependencyProperty =
            DependencyProperty.RegisterAttached(
                "Width",
                GetType(Double),
                GetType(FixedColumn))

        Private Sub New()
        End Sub

        Public Shared Function GetWidth(obj As DependencyObject) As Double
            Return CDbl(obj.GetValue(WidthProperty))
        End Function

        Public Shared Sub SetWidth(obj As DependencyObject, value As Double)
            obj.SetValue(WidthProperty, value)
        End Sub

        Public Shared Function IsFixedColumn(column As GridViewColumn) As Boolean
            If column Is Nothing Then Return False
            Return HasPropertyValue(column, WidthProperty)
        End Function

        Public Shared Function GetFixedWidth(column As GridViewColumn) As Double?
            Return GetColumnWidth(column, WidthProperty)
        End Function

        Public Shared Function ApplyWidth(column As GridViewColumn, width As Double) As GridViewColumn
            SetWidth(column, width)
            Return column
        End Function
    End Class

    ' --- Proportional columns -------------------------------------------------

    Public NotInheritable Class ProportionalColumn
        Inherits LayoutColumn

        Public Shared ReadOnly WidthProperty As DependencyProperty =
            DependencyProperty.RegisterAttached(
                "Width",
                GetType(Double),
                GetType(ProportionalColumn))

        Private Sub New()
        End Sub

        Public Shared Function GetWidth(obj As DependencyObject) As Double
            Return CDbl(obj.GetValue(WidthProperty))
        End Function

        Public Shared Sub SetWidth(obj As DependencyObject, value As Double)
            obj.SetValue(WidthProperty, value)
        End Sub

        Public Shared Function IsProportionalColumn(column As GridViewColumn) As Boolean
            If column Is Nothing Then Return False
            Return HasPropertyValue(column, WidthProperty)
        End Function

        Public Shared Function GetProportionalWidth(column As GridViewColumn) As Double?
            Return GetColumnWidth(column, WidthProperty)
        End Function
    End Class

    ' --- Range columns (min/max, optional fill) ------------------------------

    Public NotInheritable Class RangeColumn
        Inherits LayoutColumn

        Public Shared ReadOnly MinWidthProperty As DependencyProperty =
            DependencyProperty.RegisterAttached(
                "MinWidth",
                GetType(Double),
                GetType(RangeColumn))

        Public Shared ReadOnly MaxWidthProperty As DependencyProperty =
            DependencyProperty.RegisterAttached(
                "MaxWidth",
                GetType(Double),
                GetType(RangeColumn))

        Public Shared ReadOnly IsFillColumnProperty As DependencyProperty =
            DependencyProperty.RegisterAttached(
                "IsFillColumn",
                GetType(Boolean),
                GetType(RangeColumn))

        Private Sub New()
        End Sub

        Public Shared Function GetMinWidth(obj As DependencyObject) As Double
            Return CDbl(obj.GetValue(MinWidthProperty))
        End Function

        Public Shared Sub SetMinWidth(obj As DependencyObject, value As Double)
            obj.SetValue(MinWidthProperty, value)
        End Sub

        Public Shared Function GetMaxWidth(obj As DependencyObject) As Double
            Return CDbl(obj.GetValue(MaxWidthProperty))
        End Function

        Public Shared Sub SetMaxWidth(obj As DependencyObject, value As Double)
            obj.SetValue(MaxWidthProperty, value)
        End Sub

        Public Shared Function GetIsFillColumn(obj As DependencyObject) As Boolean
            Return CBool(obj.GetValue(IsFillColumnProperty))
        End Function

        Public Shared Sub SetIsFillColumn(obj As DependencyObject, value As Boolean)
            obj.SetValue(IsFillColumnProperty, value)
        End Sub

        Public Shared Function IsRangeColumn(column As GridViewColumn) As Boolean
            If column Is Nothing Then Return False
            Return HasPropertyValue(column, MinWidthProperty) OrElse HasPropertyValue(column, MaxWidthProperty) OrElse
                   HasPropertyValue(column, IsFillColumnProperty)
        End Function

        Public Shared Function GetRangeMinWidth(column As GridViewColumn) As Double?
            Return GetColumnWidth(column, MinWidthProperty)
        End Function

        Public Shared Function GetRangeMaxWidth(column As GridViewColumn) As Double?
            Return GetColumnWidth(column, MaxWidthProperty)
        End Function

        Public Shared Function GetRangeIsFillColumn(column As GridViewColumn) As Boolean?
            If column Is Nothing Then Return Nothing
            If Not HasPropertyValue(column, IsFillColumnProperty) Then Return Nothing
            Dim value = column.ReadLocalValue(IsFillColumnProperty)
            Return CBool(value)
        End Function

    End Class

End Namespace