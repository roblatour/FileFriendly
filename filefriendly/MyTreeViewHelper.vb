
Imports Microsoft.VisualBasic
Imports System
Public NotInheritable Class MyTreeViewHelper
    '
    ' The TreeViewItem that the mouse is currently directly over (or null).
    '
    Private Shared _currentItem As TreeViewItem = Nothing
    '
    ' IsMouseDirectlyOverItem: A DependencyProperty that will be true only on the 
    ' TreeViewItem that the mouse is directly over. I.e., this won't be set on that 
    ' parent item.
    '
    ' This is the only public member, and is read-only.
    '
    ' The property key (since this is a read-only DP)
    Private Shared ReadOnly IsMouseDirectlyOverItemKey As DependencyPropertyKey = DependencyProperty.RegisterAttachedReadOnly("IsMouseDirectlyOverItem", GetType(Boolean), GetType(MyTreeViewHelper), New FrameworkPropertyMetadata(Nothing, New CoerceValueCallback(AddressOf CalculateIsMouseDirectlyOverItem)))
    ' The DP itself
    Public Shared ReadOnly IsMouseDirectlyOverItemProperty As DependencyProperty = IsMouseDirectlyOverItemKey.DependencyProperty
    ' A strongly-typed getter for the property.
    Private Sub New()
    End Sub
    Public Shared Function GetIsMouseDirectlyOverItem(ByVal obj As DependencyObject) As Boolean
        Return CBool(obj.GetValue(IsMouseDirectlyOverItemProperty))
    End Function
    ' A coercion method for the property
    Private Shared Function CalculateIsMouseDirectlyOverItem(ByVal item As DependencyObject, ByVal value As Object) As Object
        ' This method is called when the IsMouseDirectlyOver property is being calculated
        ' for a TreeViewItem. 
        If item Is _currentItem Then
            Return True
        Else
            Return False
        End If
    End Function
    '
    ' UpdateOverItem: A private RoutedEvent used to find the nearest encapsulating
    ' TreeViewItem to the mouse's current position.
    '
    Private Shared ReadOnly UpdateOverItemEvent As RoutedEvent = EventManager.RegisterRoutedEvent("UpdateOverItem", RoutingStrategy.Bubble, GetType(RoutedEventHandler), GetType(MyTreeViewHelper))
    '
    ' Class constructor
    '
    Shared Sub New()
        ' Get all Mouse enter/leave events for TreeViewItem.
        EventManager.RegisterClassHandler(GetType(TreeViewItem), TreeViewItem.MouseEnterEvent, New MouseEventHandler(AddressOf OnMouseTransition), True)
        EventManager.RegisterClassHandler(GetType(TreeViewItem), TreeViewItem.MouseLeaveEvent, New MouseEventHandler(AddressOf OnMouseTransition), True)
        ' Listen for the UpdateOverItemEvent on all TreeViewItem's.
        EventManager.RegisterClassHandler(GetType(TreeViewItem), UpdateOverItemEvent, New RoutedEventHandler(AddressOf OnUpdateOverItem))
    End Sub
    '
    ' OnUpdateOverItem: This method is a listener for the UpdateOverItemEvent. When it is received,
    ' it means that the sender is the closest TreeViewItem to the mouse (closest in the sense of the
    ' tree, not geographically).
    Private Shared Sub OnUpdateOverItem(ByVal sender As Object, ByVal args As RoutedEventArgs)
        ' Mark this object as the tree view item over which the mouse
        ' is currently positioned.
        _currentItem = TryCast(sender, TreeViewItem)
        ' Tell that item to re-calculate the IsMouseDirectlyOverItem property
        _currentItem.InvalidateProperty(IsMouseDirectlyOverItemProperty)
        ' Prevent this event from notifying other tree view items higher in the tree.
        args.Handled = True
    End Sub
    '
    ' OnMouseTransition: This method is a listener for both the MouseEnter event and
    ' the MouseLeave event on TreeViewItems. It updates the _currentItem, and updates
    ' the IsMouseDirectlyOverItem property on the previous TreeViewItem and the new
    ' TreeViewItem.
    Private Shared Sub OnMouseTransition(ByVal sender As Object, ByVal args As MouseEventArgs)
        SyncLock IsMouseDirectlyOverItemProperty
            If _currentItem IsNot Nothing Then
                ' Tell the item that previously had the mouse that it no longer does.
                Dim oldItem As DependencyObject = _currentItem
                _currentItem = Nothing
                oldItem.InvalidateProperty(IsMouseDirectlyOverItemProperty)
            End If
            ' Get the element that is currently under the mouse.
            Dim currentPosition As IInputElement = Mouse.DirectlyOver
            ' See if the mouse is still over something (any element, not just a tree view item).
            If currentPosition IsNot Nothing Then
                ' Yes, the mouse is over something.
                ' Raise an event from that point. If a TreeViewItem is anywhere above this point
                ' in the tree, it will receive this event and update _currentItem.
                Dim newItemArgs As New RoutedEventArgs(UpdateOverItemEvent)
                currentPosition.RaiseEvent(newItemArgs)
            End If
        End SyncLock
    End Sub
End Class
