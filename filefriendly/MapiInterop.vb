Imports System.Runtime.InteropServices

' Minimal Extended MAPI COM interface definitions for VB.NET.
' Defines just enough of the MAPI interfaces to read folder hierarchy tables
' in batch, replacing the slow one-folder-at-a-time OOM walk.

<ComImport(), Guid("00020300-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Friend Interface IMAPISession
    <PreserveSig()> Function GetLastError(
        <[In]()> ByVal hResult As Integer,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppMAPIError As IntPtr) As Integer

    <PreserveSig()> Function GetMsgStoresTable(
        <[In]()> ByVal ulFlags As UInteger,
        <Out(), MarshalAs(UnmanagedType.Interface)> ByRef lppTable As IMAPITable) As Integer

    <PreserveSig()> Function OpenMsgStore(
        <[In]()> ByVal ulUIParam As IntPtr,
        <[In]()> ByVal cbEntryID As UInteger,
        <[In]()> ByVal lpEntryID As IntPtr,
        <[In]()> ByVal lpInterface As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out(), MarshalAs(UnmanagedType.IUnknown)> ByRef lppMDB As Object) As Integer
End Interface

<ComImport(), Guid("00020306-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Friend Interface IMsgStore
    ' IMAPIProp methods
    <PreserveSig()> Function GetLastError(
        <[In]()> ByVal hResult As Integer,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppMAPIError As IntPtr) As Integer

    <PreserveSig()> Function SaveChanges(
        <[In]()> ByVal ulFlags As UInteger) As Integer

    <PreserveSig()> Function GetProps(
        <[In]()> ByVal lpPropTagArray As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lpcValues As UInteger,
        <Out()> ByRef lppPropArray As IntPtr) As Integer

    <PreserveSig()> Function GetPropList(
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppPropTagArray As IntPtr) As Integer

    <PreserveSig()> Function OpenProperty(
        <[In]()> ByVal ulPropTag As UInteger,
        <[In]()> ByVal lpiid As IntPtr,
        <[In]()> ByVal ulInterfaceOptions As UInteger,
        <[In]()> ByVal ulFlags As UInteger,
        <Out(), MarshalAs(UnmanagedType.IUnknown)> ByRef lppUnk As Object) As Integer

    <PreserveSig()> Function SetProps(
        <[In]()> ByVal cValues As UInteger,
        <[In]()> ByVal lpPropArray As IntPtr,
        <Out()> ByRef lppProblems As IntPtr) As Integer

    <PreserveSig()> Function DeleteProps(
        <[In]()> ByVal lpPropTagArray As IntPtr,
        <Out()> ByRef lppProblems As IntPtr) As Integer

    <PreserveSig()> Function CopyTo(
        <[In]()> ByVal ciidExclude As UInteger,
        <[In]()> ByVal rgiidExclude As IntPtr,
        <[In]()> ByVal lpExcludeProps As IntPtr,
        <[In]()> ByVal ulUIParam As IntPtr,
        <[In]()> ByVal lpProgress As IntPtr,
        <[In]()> ByVal lpInterface As IntPtr,
        <[In]()> ByVal lpDestObj As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppProblems As IntPtr) As Integer

    <PreserveSig()> Function CopyProps(
        <[In]()> ByVal lpPropTagArray As IntPtr,
        <[In]()> ByVal ulUIParam As IntPtr,
        <[In]()> ByVal lpProgress As IntPtr,
        <[In]()> ByVal lpInterface As IntPtr,
        <[In]()> ByVal lpDestObj As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppProblems As IntPtr) As Integer

    <PreserveSig()> Function GetNamesFromIDs(
        <[In](), Out()> ByRef lppPropTags As IntPtr,
        <[In]()> ByVal lpPropSetGuid As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lpcNames As UInteger,
        <Out()> ByRef lpppNames As IntPtr) As Integer

    <PreserveSig()> Function GetIDsFromNames(
        <[In]()> ByVal cNames As UInteger,
        <[In]()> ByVal lppPropNames As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppPropTags As IntPtr) As Integer

    ' IMsgStore-specific methods
    <PreserveSig()> Function Advise(
        <[In]()> ByVal cbEntryID As UInteger,
        <[In]()> ByVal lpEntryID As IntPtr,
        <[In]()> ByVal ulEventMask As UInteger,
        <[In]()> ByVal lpAdviseSink As IntPtr,
        <Out()> ByRef lpulConnection As UInteger) As Integer

    <PreserveSig()> Function Unadvise(
        <[In]()> ByVal ulConnection As UInteger) As Integer

    <PreserveSig()> Function CompareEntryIDs(
        <[In]()> ByVal cbEntryID1 As UInteger,
        <[In]()> ByVal lpEntryID1 As IntPtr,
        <[In]()> ByVal cbEntryID2 As UInteger,
        <[In]()> ByVal lpEntryID2 As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lpulResult As UInteger) As Integer

    <PreserveSig()> Function OpenEntry(
        <[In]()> ByVal cbEntryID As UInteger,
        <[In]()> ByVal lpEntryID As IntPtr,
        <[In]()> ByVal lpInterface As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lpulObjType As UInteger,
        <Out(), MarshalAs(UnmanagedType.IUnknown)> ByRef lppUnk As Object) As Integer
End Interface

<ComImport(), Guid("0002030B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Friend Interface IMAPIContainer
    ' IMAPIProp methods
    <PreserveSig()> Function GetLastError(
        <[In]()> ByVal hResult As Integer,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppMAPIError As IntPtr) As Integer

    <PreserveSig()> Function SaveChanges(
        <[In]()> ByVal ulFlags As UInteger) As Integer

    <PreserveSig()> Function GetProps(
        <[In]()> ByVal lpPropTagArray As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lpcValues As UInteger,
        <Out()> ByRef lppPropArray As IntPtr) As Integer

    <PreserveSig()> Function GetPropList(
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppPropTagArray As IntPtr) As Integer

    <PreserveSig()> Function OpenProperty(
        <[In]()> ByVal ulPropTag As UInteger,
        <[In]()> ByVal lpiid As IntPtr,
        <[In]()> ByVal ulInterfaceOptions As UInteger,
        <[In]()> ByVal ulFlags As UInteger,
        <Out(), MarshalAs(UnmanagedType.IUnknown)> ByRef lppUnk As Object) As Integer

    <PreserveSig()> Function SetProps(
        <[In]()> ByVal cValues As UInteger,
        <[In]()> ByVal lpPropArray As IntPtr,
        <Out()> ByRef lppProblems As IntPtr) As Integer

    <PreserveSig()> Function DeleteProps(
        <[In]()> ByVal lpPropTagArray As IntPtr,
        <Out()> ByRef lppProblems As IntPtr) As Integer

    <PreserveSig()> Function CopyTo(
        <[In]()> ByVal ciidExclude As UInteger,
        <[In]()> ByVal rgiidExclude As IntPtr,
        <[In]()> ByVal lpExcludeProps As IntPtr,
        <[In]()> ByVal ulUIParam As IntPtr,
        <[In]()> ByVal lpProgress As IntPtr,
        <[In]()> ByVal lpInterface As IntPtr,
        <[In]()> ByVal lpDestObj As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppProblems As IntPtr) As Integer

    <PreserveSig()> Function CopyProps(
        <[In]()> ByVal lpPropTagArray As IntPtr,
        <[In]()> ByVal ulUIParam As IntPtr,
        <[In]()> ByVal lpProgress As IntPtr,
        <[In]()> ByVal lpInterface As IntPtr,
        <[In]()> ByVal lpDestObj As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppProblems As IntPtr) As Integer

    <PreserveSig()> Function GetNamesFromIDs(
        <[In](), Out()> ByRef lppPropTags As IntPtr,
        <[In]()> ByVal lpPropSetGuid As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lpcNames As UInteger,
        <Out()> ByRef lpppNames As IntPtr) As Integer

    <PreserveSig()> Function GetIDsFromNames(
        <[In]()> ByVal cNames As UInteger,
        <[In]()> ByVal lppPropNames As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppPropTags As IntPtr) As Integer

    ' IMAPIContainer-specific methods
    <PreserveSig()> Function GetContentsTable(
        <[In]()> ByVal ulFlags As UInteger,
        <Out(), MarshalAs(UnmanagedType.Interface)> ByRef lppTable As IMAPITable) As Integer

    <PreserveSig()> Function GetHierarchyTable(
        <[In]()> ByVal ulFlags As UInteger,
        <Out(), MarshalAs(UnmanagedType.Interface)> ByRef lppTable As IMAPITable) As Integer

    <PreserveSig()> Function OpenEntry(
        <[In]()> ByVal cbEntryID As UInteger,
        <[In]()> ByVal lpEntryID As IntPtr,
        <[In]()> ByVal lpInterface As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lpulObjType As UInteger,
        <Out(), MarshalAs(UnmanagedType.IUnknown)> ByRef lppUnk As Object) As Integer

    <PreserveSig()> Function SetSearchCriteria(
        <[In]()> ByVal lpContainerTags As IntPtr,
        <[In]()> ByVal ulSearchFlags As UInteger) As Integer

    <PreserveSig()> Function GetSearchCriteria(
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppContainerTags As IntPtr,
        <Out()> ByRef lppSearchState As UInteger) As Integer
End Interface

<ComImport(), Guid("00020312-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Friend Interface IMAPIFolder
    ' IMAPIProp methods
    <PreserveSig()> Function GetLastError(
        <[In]()> ByVal hResult As Integer,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppMAPIError As IntPtr) As Integer

    <PreserveSig()> Function SaveChanges(
        <[In]()> ByVal ulFlags As UInteger) As Integer

    <PreserveSig()> Function GetProps(
        <[In]()> ByVal lpPropTagArray As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lpcValues As UInteger,
        <Out()> ByRef lppPropArray As IntPtr) As Integer

    <PreserveSig()> Function GetPropList(
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppPropTagArray As IntPtr) As Integer

    <PreserveSig()> Function OpenProperty(
        <[In]()> ByVal ulPropTag As UInteger,
        <[In]()> ByVal lpiid As IntPtr,
        <[In]()> ByVal ulInterfaceOptions As UInteger,
        <[In]()> ByVal ulFlags As UInteger,
        <Out(), MarshalAs(UnmanagedType.IUnknown)> ByRef lppUnk As Object) As Integer

    <PreserveSig()> Function SetProps(
        <[In]()> ByVal cValues As UInteger,
        <[In]()> ByVal lpPropArray As IntPtr,
        <Out()> ByRef lppProblems As IntPtr) As Integer

    <PreserveSig()> Function DeleteProps(
        <[In]()> ByVal lpPropTagArray As IntPtr,
        <Out()> ByRef lppProblems As IntPtr) As Integer

    <PreserveSig()> Function CopyTo(
        <[In]()> ByVal ciidExclude As UInteger,
        <[In]()> ByVal rgiidExclude As IntPtr,
        <[In]()> ByVal lpExcludeProps As IntPtr,
        <[In]()> ByVal ulUIParam As IntPtr,
        <[In]()> ByVal lpProgress As IntPtr,
        <[In]()> ByVal lpInterface As IntPtr,
        <[In]()> ByVal lpDestObj As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppProblems As IntPtr) As Integer

    <PreserveSig()> Function CopyProps(
        <[In]()> ByVal lpPropTagArray As IntPtr,
        <[In]()> ByVal ulUIParam As IntPtr,
        <[In]()> ByVal lpProgress As IntPtr,
        <[In]()> ByVal lpInterface As IntPtr,
        <[In]()> ByVal lpDestObj As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppProblems As IntPtr) As Integer

    <PreserveSig()> Function GetNamesFromIDs(
        <[In](), Out()> ByRef lppPropTags As IntPtr,
        <[In]()> ByVal lpPropSetGuid As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lpcNames As UInteger,
        <Out()> ByRef lpppNames As IntPtr) As Integer

    <PreserveSig()> Function GetIDsFromNames(
        <[In]()> ByVal cNames As UInteger,
        <[In]()> ByVal lppPropNames As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppPropTags As IntPtr) As Integer

    ' IMAPIContainer-specific methods
    <PreserveSig()> Function GetContentsTable(
        <[In]()> ByVal ulFlags As UInteger,
        <Out(), MarshalAs(UnmanagedType.Interface)> ByRef lppTable As IMAPITable) As Integer

    <PreserveSig()> Function GetHierarchyTable(
        <[In]()> ByVal ulFlags As UInteger,
        <Out(), MarshalAs(UnmanagedType.Interface)> ByRef lppTable As IMAPITable) As Integer

    <PreserveSig()> Function OpenEntry(
        <[In]()> ByVal cbEntryID As UInteger,
        <[In]()> ByVal lpEntryID As IntPtr,
        <[In]()> ByVal lpInterface As IntPtr,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lpulObjType As UInteger,
        <Out(), MarshalAs(UnmanagedType.IUnknown)> ByRef lppUnk As Object) As Integer

    <PreserveSig()> Function SetSearchCriteria(
        <[In]()> ByVal lpContainerTags As IntPtr,
        <[In]()> ByVal ulSearchFlags As UInteger) As Integer

    <PreserveSig()> Function GetSearchCriteria(
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppContainerTags As IntPtr,
        <Out()> ByRef lppSearchState As UInteger) As Integer
End Interface

<ComImport(), Guid("00020307-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Friend Interface IMAPITable
    <PreserveSig()> Function GetLastError(
        <[In]()> ByVal hResult As Integer,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppMAPIError As IntPtr) As Integer

    <PreserveSig()> Function Advise(
        <[In]()> ByVal ulEventMask As UInteger,
        <[In]()> ByVal lpAdviseSink As IntPtr,
        <Out()> ByRef lpulConnection As UInteger) As Integer

    <PreserveSig()> Function Unadvise(
        <[In]()> ByVal ulConnection As UInteger) As Integer

    <PreserveSig()> Function GetStatus(
        <Out()> ByRef lpulTableStatus As UInteger,
        <Out()> ByRef lpulTableType As UInteger) As Integer

    <PreserveSig()> Function SetColumns(
        <[In]()> ByVal lpPropTagArray As IntPtr,
        <[In]()> ByVal ulFlags As UInteger) As Integer

    <PreserveSig()> Function QueryColumns(
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppPropTagArray As IntPtr) As Integer

    <PreserveSig()> Function GetRowCount(
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lpcRowCount As UInteger) As Integer

    <PreserveSig()> Function SeekRow(
        <[In]()> ByVal bkOrigin As UInteger,
        <[In]()> ByVal lRowCount As Integer,
        <Out()> ByRef lplRowsSought As Integer) As Integer

    <PreserveSig()> Function SeekRowApprox(
        <[In]()> ByVal ulNumerator As UInteger,
        <[In]()> ByVal ulDenominator As UInteger) As Integer

    <PreserveSig()> Function QueryRows(
        <[In]()> ByVal lRowCount As Integer,
        <[In]()> ByVal ulFlags As UInteger,
        <Out()> ByRef lppRows As IntPtr) As Integer

    <PreserveSig()> Function Abort() As Integer
End Interface

' MAPI property types
Friend Module MapiPropTypes
    Friend Const PT_LONG As UInteger = &H3
    Friend Const PT_BOOLEAN As UInteger = &HB
    Friend Const PT_STRING8 As UInteger = &H1E
    Friend Const PT_UNICODE As UInteger = &H1F
    Friend Const PT_BINARY As UInteger = &H102
    Friend Const PT_ERROR As UInteger = &HA
    Friend Const PT_NULL As UInteger = &H1
End Module

' MAPI property tags: PROP_TAG(type, id) = (id << 16) | type
Friend Module MapiPropTags
    Friend Const PR_DISPLAY_NAME_W As UInteger = &H3001001FUI
    Friend Const PR_DISPLAY_NAME_A As UInteger = &H3001001EUI
    Friend Const PR_ENTRYID As UInteger = &H0FFF0102UI
    Friend Const PR_DEPTH As UInteger = &H30050003UI
    Friend Const PR_FOLDER_TYPE As UInteger = &H36010003UI
    Friend Const PR_SUBFOLDERS As UInteger = &H360A000BUI
    Friend Const PR_CONTENT_COUNT As UInteger = &H36020003UI
    Friend Const PR_CONTENT_UNREAD As UInteger = &H36030003UI
    Friend Const PR_CONTAINER_CLASS_W As UInteger = &H3613001FUI
    Friend Const PR_CONTAINER_CLASS_A As UInteger = &H3613001EUI
    Friend Const PR_STORE_ENTRYID As UInteger = &H0FFB0102UI
    Friend Const PR_IPM_SUBTREE_ENTRYID As UInteger = &H35E00102UI
End Module

' MAPI flags
Friend Module MapiFlags
    Friend Const MAPI_BEST_ACCESS As UInteger = &H10UI
    Friend Const MAPI_DEFERRED_ERRORS As UInteger = &H8UI
    Friend Const MAPI_EXTENDED As UInteger = &H20UI
    Friend Const MAPI_USE_DEFAULT As UInteger = &H40UI
    Friend Const MAPI_UNICODE As UInteger = &H80000000UI
    Friend Const TABLE_START As UInteger = &HFFFFFFFEUI
End Module

Friend Module MapiNative
    <DllImport("mapi32.dll", CharSet:=CharSet.Unicode)>
    Friend Function MAPIInitialize(ByVal lpMapiInit As IntPtr) As Integer
    End Function

    <DllImport("mapi32.dll", CharSet:=CharSet.Unicode)>
    Friend Function MAPILogonEx(ByVal ulUIParam As IntPtr,
                                ByVal lpszProfileName As String,
                                ByVal lpszPassword As String,
                                ByVal flFlags As UInteger,
                                ByRef lppSession As IntPtr) As Integer
    End Function

    <DllImport("mapi32.dll")>
    Friend Sub MAPIUninitialize()
    End Sub
End Module

' MAPI folder types (PR_FOLDER_TYPE values)
Friend Module MapiFolderTypes
    Friend Const FOLDER_GENERIC As UInteger = 1UI
    Friend Const FOLDER_SEARCH As UInteger = 2UI
End Module

' Helper class for reading MAPI structures from unmanaged memory
Friend Module MapiHelpers

    <DllImport("mapi32.dll", CharSet:=CharSet.Unicode)>
    Friend Function MAPIFreeBuffer(ByVal lpBuffer As IntPtr) As Integer
    End Function

    ' Size of SPropValue structure depends on pointer size
    ' 32-bit: 4(ulPropTag) + 4(dwAlignPad) + 8(union) = 16
    ' 64-bit: 4(ulPropTag) + 4(dwAlignPad) + 16(union) = 24
    Friend ReadOnly SPropValueSize As Integer = If(IntPtr.Size = 8, 24, 16)

    ' Offset of SBinary.lpb within SPropValue
    ' 32-bit: 4 + 4 + 4 = 12
    ' 64-bit: 4 + 4 + 4 + 4(pad) = 16
    Friend ReadOnly SBinaryLpbOffset As Integer = If(IntPtr.Size = 8, 16, 12)

    ' Size of SRow structure
    ' 32-bit: 4(ulAdrEntryPad) + 4(cValues) + 4(lpProps) = 12
    ' 64-bit: 4(ulAdrEntryPad) + 4(cValues) + 8(lpProps) = 16
    Friend ReadOnly SRowSize As Integer = If(IntPtr.Size = 8, 16, 12)

    ' Offset of first SRow within SRowSet (after cRows ULONG + alignment padding)
    ' 32-bit: cRows(4) + aRow at 4 (no padding, SRow align = 4)
    ' 64-bit: cRows(4) + 4 bytes padding + aRow at 8 (SRow align = 8)
    Friend ReadOnly SRowSetFirstRowOffset As Integer = If(IntPtr.Size = 8, 8, 4)

    ' Allocates an SPropTagArray in unmanaged memory
    Friend Function AllocPropTagArray(ByVal tags As UInteger()) As IntPtr
        Dim byteSize As Integer = 4 * (1 + tags.Length)
        Dim ptr As IntPtr = Marshal.AllocHGlobal(byteSize)
        Marshal.WriteInt32(ptr, 0, tags.Length)
        For i As Integer = 0 To tags.Length - 1
            Marshal.WriteInt32(ptr, 4 * (1 + i), CInt(tags(i)))
        Next
        Return ptr
    End Function

    ' Reads a LONG property value from an SPropValue array
    Friend Function GetLongProperty(ByVal propArray As IntPtr, ByVal cValues As UInteger, ByVal propTag As UInteger) As Integer
        For i As Integer = 0 To CInt(cValues) - 1
            Dim propPtr As IntPtr = IntPtr.Add(propArray, i * SPropValueSize)
            Dim tag As UInteger = CUInt(Marshal.ReadInt32(propPtr, 0))
            If tag = propTag Then
                Return Marshal.ReadInt32(propPtr, 8)
            End If
        Next
        Return 0
    End Function

    ' Reads a BOOLEAN property value from an SPropValue array
    Friend Function GetBoolProperty(ByVal propArray As IntPtr, ByVal cValues As UInteger, ByVal propTag As UInteger) As Boolean
        For i As Integer = 0 To CInt(cValues) - 1
            Dim propPtr As IntPtr = IntPtr.Add(propArray, i * SPropValueSize)
            Dim tag As UInteger = CUInt(Marshal.ReadInt32(propPtr, 0))
            If tag = propTag Then
                Return Marshal.ReadInt16(propPtr, 8) <> 0
            End If
        Next
        Return False
    End Function

    ' Reads a UNICODE string property value from an SPropValue array.
    ' Returns Nothing if the property is absent or an error.
    Friend Function GetStringProperty(ByVal propArray As IntPtr, ByVal cValues As UInteger, ByVal propTag As UInteger) As String
        For i As Integer = 0 To CInt(cValues) - 1
            Dim propPtr As IntPtr = IntPtr.Add(propArray, i * SPropValueSize)
            Dim tag As UInteger = CUInt(Marshal.ReadInt32(propPtr, 0))
            If tag = propTag Then
                Dim strPtr As IntPtr = Marshal.ReadIntPtr(propPtr, 8)
                If strPtr <> IntPtr.Zero Then
                    Return Marshal.PtrToStringUni(strPtr)
                End If
                Return Nothing
            End If
            ' Fall back to ANSI variant if Unicode tag not found
            If propTag = MapiPropTags.PR_DISPLAY_NAME_W AndAlso tag = MapiPropTags.PR_DISPLAY_NAME_A Then
                Dim strPtr As IntPtr = Marshal.ReadIntPtr(propPtr, 8)
                If strPtr <> IntPtr.Zero Then
                    Return Marshal.PtrToStringAnsi(strPtr)
                End If
                Return Nothing
            End If
            If propTag = MapiPropTags.PR_CONTAINER_CLASS_W AndAlso tag = MapiPropTags.PR_CONTAINER_CLASS_A Then
                Dim strPtr As IntPtr = Marshal.ReadIntPtr(propPtr, 8)
                If strPtr <> IntPtr.Zero Then
                    Return Marshal.PtrToStringAnsi(strPtr)
                End If
                Return Nothing
            End If
        Next
        Return Nothing
    End Function

    ' Reads a BINARY property value from an SPropValue array and returns it as a hex string.
    Friend Function GetBinaryPropertyAsHex(ByVal propArray As IntPtr, ByVal cValues As UInteger, ByVal propTag As UInteger) As String
        For i As Integer = 0 To CInt(cValues) - 1
            Dim propPtr As IntPtr = IntPtr.Add(propArray, i * SPropValueSize)
            Dim tag As UInteger = CUInt(Marshal.ReadInt32(propPtr, 0))
            If tag = propTag Then
                Dim cb As Integer = Marshal.ReadInt32(propPtr, 8)
                Dim lpb As IntPtr = Marshal.ReadIntPtr(propPtr, SBinaryLpbOffset)
                If cb > 0 AndAlso lpb <> IntPtr.Zero Then
                    Dim bytes(cb - 1) As Byte
                    Marshal.Copy(lpb, bytes, 0, cb)
                    Return BitConverter.ToString(bytes).Replace("-", "")
                End If
                Return ""
            End If
        Next
        Return Nothing
    End Function

    ' Reads a BINARY property as a raw byte array.
    Friend Function GetBinaryProperty(ByVal propArray As IntPtr, ByVal cValues As UInteger, ByVal propTag As UInteger, ByRef cb As Integer) As IntPtr
        For i As Integer = 0 To CInt(cValues) - 1
            Dim propPtr As IntPtr = IntPtr.Add(propArray, i * SPropValueSize)
            Dim tag As UInteger = CUInt(Marshal.ReadInt32(propPtr, 0))
            If tag = propTag Then
                cb = Marshal.ReadInt32(propPtr, 8)
                Return Marshal.ReadIntPtr(propPtr, SBinaryLpbOffset)
            End If
        Next
        cb = 0
        Return IntPtr.Zero
    End Function

    ' Checks if a property tag in an SPropValue array has PT_ERROR type
    Friend Function IsPropertyError(ByVal propArray As IntPtr, ByVal cValues As UInteger, ByVal propTag As UInteger) As Boolean
        For i As Integer = 0 To CInt(cValues) - 1
            Dim propPtr As IntPtr = IntPtr.Add(propArray, i * SPropValueSize)
            Dim tag As UInteger = CUInt(Marshal.ReadInt32(propPtr, 0))
            If (tag And &HFFFF0000UI) = (propTag And &HFFFF0000UI) Then
                Dim typePart As UInteger = tag And &HFFFFUI
                Return typePart = MapiPropTypes.PT_ERROR
            End If
        Next
        Return True
    End Function

    ' Converts a byte array to a hex string (for EntryID/StoreID)
    Friend Function BytesToHex(ByVal bytes As Byte()) As String
        Return BitConverter.ToString(bytes).Replace("-", "")
    End Function

    ' Determines if a container class string indicates a mail folder
    Friend Function IsMailFolder(ByVal containerClass As String) As Boolean
        If String.IsNullOrEmpty(containerClass) Then Return False
        Return containerClass.StartsWith("IPF.Note", StringComparison.OrdinalIgnoreCase) OrElse
               containerClass.StartsWith("IPF.Imap", StringComparison.OrdinalIgnoreCase)
    End Function

End Module

' =====================================================================
' Raw vtable call infrastructure
' MAPI interfaces are not registered in the Windows registry, so the
' CLR's ComImport QueryInterface fails with REGDB_E_CLASSNOTREG.
' We bypass COM interop and call vtable function pointers directly
' via delegates.
' =====================================================================

' Delegate definitions for each MAPI method we need.
' The first parameter (pThis) is always the COM interface pointer.

<UnmanagedFunctionPointer(CallingConvention.StdCall)>
Friend Delegate Function MapiReleaseDelegate(ByVal pThis As IntPtr) As Integer

' IMAPISession vtable indices (after IUnknown [0-2]):
'   [3] GetLastError  [4] GetMsgStoresTable  [5] OpenMsgStore
<UnmanagedFunctionPointer(CallingConvention.StdCall)>
Friend Delegate Function MapiGetMsgStoresTableDlg(
    ByVal pThis As IntPtr,
    ByVal ulFlags As UInteger,
    ByRef lppTable As IntPtr) As Integer

<UnmanagedFunctionPointer(CallingConvention.StdCall)>
Friend Delegate Function MapiOpenMsgStoreDlg(
    ByVal pThis As IntPtr,
    ByVal ulUIParam As IntPtr,
    ByVal cbEntryID As UInteger,
    ByVal lpEntryID As IntPtr,
    ByVal lpInterface As IntPtr,
    ByVal ulFlags As UInteger,
    ByRef lppMDB As IntPtr) As Integer

<UnmanagedFunctionPointer(CallingConvention.StdCall)>
Friend Delegate Function MapiLogoffDlg(
    ByVal pThis As IntPtr,
    ByVal ulUIParam As IntPtr,
    ByVal ulFlags As UInteger,
    ByVal ulReserved As UInteger) As Integer

' IMsgStore vtable indices (after IUnknown + IMAPIProp):
'   [17] OpenEntry
' IMAPIFolder vtable indices (after IUnknown + IMAPIProp):
'   [5]  GetProps  [15] GetHierarchyTable  [16] OpenEntry
<UnmanagedFunctionPointer(CallingConvention.StdCall)>
Friend Delegate Function MapiGetPropsDlg(
    ByVal pThis As IntPtr,
    ByVal lpPropTagArray As IntPtr,
    ByVal ulFlags As UInteger,
    ByRef lpcValues As UInteger,
    ByRef lppPropArray As IntPtr) As Integer

<UnmanagedFunctionPointer(CallingConvention.StdCall)>
Friend Delegate Function MapiGetHierarchyTableDlg(
    ByVal pThis As IntPtr,
    ByVal ulFlags As UInteger,
    ByRef lppTable As IntPtr) As Integer

<UnmanagedFunctionPointer(CallingConvention.StdCall)>
Friend Delegate Function MapiOpenEntryDlg(
    ByVal pThis As IntPtr,
    ByVal cbEntryID As UInteger,
    ByVal lpEntryID As IntPtr,
    ByVal lpInterface As IntPtr,
    ByVal ulFlags As UInteger,
    ByRef lpulObjType As UInteger,
    ByRef lppUnk As IntPtr) As Integer

' IMAPITable vtable indices (after IUnknown [0-2]):
'   [7] SetColumns  [19] QueryRows
<UnmanagedFunctionPointer(CallingConvention.StdCall)>
Friend Delegate Function MapiSetColumnsDlg(
    ByVal pThis As IntPtr,
    ByVal lpPropTagArray As IntPtr,
    ByVal ulFlags As UInteger) As Integer

<UnmanagedFunctionPointer(CallingConvention.StdCall)>
Friend Delegate Function MapiQueryRowsDlg(
    ByVal pThis As IntPtr,
    ByVal lRowCount As Integer,
    ByVal ulFlags As UInteger,
    ByRef lppRows As IntPtr) As Integer

' Helper module for raw vtable access
Friend Module MapiVtable

    ' Vtable indices for each interface/method we use
    Friend Const IdxRelease As Integer = 2          ' IUnknown::Release
    Friend Const IdxGetMsgStoresTable As Integer = 4  ' IMAPISession
    Friend Const IdxOpenMsgStore As Integer = 5       ' IMAPISession
    Friend Const IdxGetProps As Integer = 5           ' IMAPIProp (IMsgStore, IMAPIFolder)
    Friend Const IdxLogoff As Integer = 17            ' IMAPISession
    Friend Const IdxOpenEntryMsgStore As Integer = 17 ' IMsgStore
    Friend Const IdxGetHierarchyTable As Integer = 15 ' IMAPIContainer (IMAPIFolder)
    Friend Const IdxOpenEntryFolder As Integer = 16   ' IMAPIContainer (IMAPIFolder)
    Friend Const IdxSetColumns As Integer = 7         ' IMAPITable
    Friend Const IdxQueryRows As Integer = 19         ' IMAPITable

    ' Reads a function pointer from a COM object's vtable
    Friend Function GetMethod(ByVal pUnk As IntPtr, ByVal vtableIndex As Integer) As IntPtr
        Dim vtable As IntPtr = Marshal.ReadIntPtr(pUnk)
        Return Marshal.ReadIntPtr(vtable, vtableIndex * IntPtr.Size)
    End Function

    ' Calls IUnknown::Release on a raw COM pointer
    Friend Sub Release(ByVal pUnk As IntPtr)
        If pUnk = IntPtr.Zero Then Return
        Dim fn As IntPtr = GetMethod(pUnk, IdxRelease)
        Dim dlg As MapiReleaseDelegate = Marshal.GetDelegateForFunctionPointer(fn, GetType(MapiReleaseDelegate))
        dlg(pUnk)
    End Sub

    Friend Function Logoff(ByVal pSession As IntPtr) As Integer
        Dim fn As IntPtr = GetMethod(pSession, IdxLogoff)
        Dim dlg As MapiLogoffDlg = Marshal.GetDelegateForFunctionPointer(fn, GetType(MapiLogoffDlg))
        Return dlg(pSession, IntPtr.Zero, 0, 0)
    End Function

    ' IMAPISession::GetMsgStoresTable
    Friend Function GetMsgStoresTable(ByVal pSession As IntPtr, ByVal ulFlags As UInteger, ByRef lppTable As IntPtr) As Integer
        Dim fn As IntPtr = GetMethod(pSession, IdxGetMsgStoresTable)
        Dim dlg As MapiGetMsgStoresTableDlg = Marshal.GetDelegateForFunctionPointer(fn, GetType(MapiGetMsgStoresTableDlg))
        Return dlg(pSession, ulFlags, lppTable)
    End Function

    ' IMAPISession::OpenMsgStore
    Friend Function OpenMsgStore(ByVal pSession As IntPtr, ByVal ulUIParam As IntPtr,
                                  ByVal cbEntryID As UInteger, ByVal lpEntryID As IntPtr,
                                  ByVal lpInterface As IntPtr, ByVal ulFlags As UInteger,
                                  ByRef lppMDB As IntPtr) As Integer
        Dim fn As IntPtr = GetMethod(pSession, IdxOpenMsgStore)
        Dim dlg As MapiOpenMsgStoreDlg = Marshal.GetDelegateForFunctionPointer(fn, GetType(MapiOpenMsgStoreDlg))
        Return dlg(pSession, ulUIParam, cbEntryID, lpEntryID, lpInterface, ulFlags, lppMDB)
    End Function

    ' IMsgStore::OpenEntry
    Friend Function MsgStoreOpenEntry(ByVal pStore As IntPtr,
                                       ByVal cbEntryID As UInteger, ByVal lpEntryID As IntPtr,
                                       ByVal lpInterface As IntPtr, ByVal ulFlags As UInteger,
                                       ByRef lpulObjType As UInteger, ByRef lppUnk As IntPtr) As Integer
        Dim fn As IntPtr = GetMethod(pStore, IdxOpenEntryMsgStore)
        Dim dlg As MapiOpenEntryDlg = Marshal.GetDelegateForFunctionPointer(fn, GetType(MapiOpenEntryDlg))
        Return dlg(pStore, cbEntryID, lpEntryID, lpInterface, ulFlags, lpulObjType, lppUnk)
    End Function

    ' IMAPIFolder::GetProps
    Friend Function FolderGetProps(ByVal pFolder As IntPtr,
                                    ByVal lpPropTagArray As IntPtr, ByVal ulFlags As UInteger,
                                    ByRef lpcValues As UInteger, ByRef lppPropArray As IntPtr) As Integer
        Dim fn As IntPtr = GetMethod(pFolder, IdxGetProps)
        Dim dlg As MapiGetPropsDlg = Marshal.GetDelegateForFunctionPointer(fn, GetType(MapiGetPropsDlg))
        Return dlg(pFolder, lpPropTagArray, ulFlags, lpcValues, lppPropArray)
    End Function

    ' IMAPIFolder::GetHierarchyTable
    Friend Function GetHierarchyTable(ByVal pFolder As IntPtr, ByVal ulFlags As UInteger,
                                       ByRef lppTable As IntPtr) As Integer
        Dim fn As IntPtr = GetMethod(pFolder, IdxGetHierarchyTable)
        Dim dlg As MapiGetHierarchyTableDlg = Marshal.GetDelegateForFunctionPointer(fn, GetType(MapiGetHierarchyTableDlg))
        Return dlg(pFolder, ulFlags, lppTable)
    End Function

    ' IMAPIFolder::OpenEntry
    Friend Function FolderOpenEntry(ByVal pFolder As IntPtr,
                                     ByVal cbEntryID As UInteger, ByVal lpEntryID As IntPtr,
                                     ByVal lpInterface As IntPtr, ByVal ulFlags As UInteger,
                                     ByRef lpulObjType As UInteger, ByRef lppUnk As IntPtr) As Integer
        Dim fn As IntPtr = GetMethod(pFolder, IdxOpenEntryFolder)
        Dim dlg As MapiOpenEntryDlg = Marshal.GetDelegateForFunctionPointer(fn, GetType(MapiOpenEntryDlg))
        Return dlg(pFolder, cbEntryID, lpEntryID, lpInterface, ulFlags, lpulObjType, lppUnk)
    End Function

    ' IMAPITable::SetColumns
    Friend Function SetColumns(ByVal pTable As IntPtr,
                                ByVal lpPropTagArray As IntPtr, ByVal ulFlags As UInteger) As Integer
        Dim fn As IntPtr = GetMethod(pTable, IdxSetColumns)
        Dim dlg As MapiSetColumnsDlg = Marshal.GetDelegateForFunctionPointer(fn, GetType(MapiSetColumnsDlg))
        Return dlg(pTable, lpPropTagArray, ulFlags)
    End Function

    ' IMAPITable::QueryRows
    Friend Function QueryRows(ByVal pTable As IntPtr,
                               ByVal lRowCount As Integer, ByVal ulFlags As UInteger,
                               ByRef lppRows As IntPtr) As Integer
        Dim fn As IntPtr = GetMethod(pTable, IdxQueryRows)
        Dim dlg As MapiQueryRowsDlg = Marshal.GetDelegateForFunctionPointer(fn, GetType(MapiQueryRowsDlg))
        Return dlg(pTable, lRowCount, ulFlags, lppRows)
    End Function

End Module
