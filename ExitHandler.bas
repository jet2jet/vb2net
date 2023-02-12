Attribute VB_Name = "ExitHandler"
' Copyright (C) 2023 jet
' For more information about license, see LICENSE.
'
' Helper module for handling on application-exit
Option Explicit

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' Instance data of 'MyClass'
Private Type MyClassData
    vtblPtr As LongPtr
    RefCount As Long
#If Win64 Then ' Whether the platform is x64
    Padding As Long
#End If
End Type

' The data of virtual function table
Private Type IUnknownVtbl
    QueryInterface As LongPtr
    AddRef As LongPtr
    Release As LongPtr
End Type

Private Const S_OK As Long = 0
Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_POINTER As Long = &H80004003

Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
    (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)
Public Declare PtrSafe Function GetProcessHeap Lib "kernel32.dll" () As LongPtr
Public Declare PtrSafe Function HeapAlloc Lib "kernel32.dll" _
    (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Public Declare PtrSafe Function HeapFree Lib "kernel32.dll" _
    (ByVal hHeap As LongPtr, ByVal dwFlags As Long, ByVal lpMem As LongPtr) As Boolean

Public Declare PtrSafe Function CoTaskMemAlloc Lib "ole32.dll" _
    (ByVal cb As LongPtr) As LongPtr
Public Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" _
    (ByVal pv As LongPtr)
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" _
    (ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Integer, _
    ByRef prgpvarg As LongPtr, _
    ByRef pvargResult As Variant) As Long

' The MyClass instance kept during running VBA
Dim m_unk As IUnknown
Dim m_collHandlers As Collection

' Helper function to get function address
Private Function GetAddressOf(ByVal func As LongPtr) As LongPtr
    GetAddressOf = func
End Function

' Returns the pointer referring the data block including MyClassData and IUnknownVtbl
Private Function CreateInstanceMemory() As LongPtr
    Dim p As LongPtr, d As MyClassData, v As IUnknownVtbl
    ' allocate the size of sum of MyClassData and IUnknownVtbl
    p = CoTaskMemAlloc(Len(d) + Len(v))
    If p <> 0 Then
        ' always set 1 for first reference count
        d.RefCount = 1
        ' set the address of (p + sizeof(MyClassData)) due to putting IUnknownVtbl just after MyClassData
        d.vtblPtr = p + Len(d)
        ' fill p by MyClassData
        Call CopyMemory(ByVal p, d, Len(d))
        ' create virtual function table
        v.QueryInterface = GetAddressOf(AddressOf My_QueryInterface)
        v.AddRef = GetAddressOf(AddressOf My_AddRef)
        v.Release = GetAddressOf(AddressOf My_Release)
        ' copy virtual function table into (p + Len(d))
        Call CopyMemory(ByVal d.vtblPtr, v, Len(v))
    End If
    CreateInstanceMemory = p
End Function

' HRESULT STDMETHODCALLTYPE QueryInterface(THIS_ REFIID refiid, LPVOID FAR* ppv)
' Called when requested to other interface
' (defined ppv as ByVal to check whether nullptr)
Private Function My_QueryInterface(ByVal This As LongPtr, ByRef refiid As IID, ByVal ppv As LongPtr) As Long
    Debug.Print "My_QueryInterface"
    If ppv = 0 Then
        My_QueryInterface = E_POINTER
        Exit Function
    End If
    ' check whether refiid refers to IID_IUnknown: {00000000-0000-0000-C000-000000000046}
    If refiid.Data1 = 0 And refiid.Data2 = 0 And refiid.Data3 = 0 And _
        refiid.Data4(0) = &HC0 And refiid.Data4(1) = 0 And _
        refiid.Data4(2) = 0 And refiid.Data4(3) = 0 And _
        refiid.Data4(4) = 0 And refiid.Data4(5) = 0 And _
        refiid.Data4(6) = 0 And refiid.Data4(7) = &H46& Then
        ' if IID_IUnknown, copy the address of This (the value of This) into the pointer ppv
        Call CopyMemory(ByVal ppv, This, Len(This))
        ' increment reference count
        Call My_AddRef(This)
        My_QueryInterface = S_OK
        Exit Function
    End If
    ' do not support interfaces other than IID_IUnknown
    My_QueryInterface = E_NOINTERFACE
End Function

' ULONG STDMETHODCALLTYPE AddRef(THIS)
' Called when incrementing reference count
Private Function My_AddRef(ByVal This As LongPtr) As Long
    Dim d As MyClassData
    ' copy the instance data into d first,
    ' and restore to d after incrementing reference count
    Call CopyMemory(d, ByVal This, Len(d))
    d.RefCount = d.RefCount + 1
    Call CopyMemory(ByVal This, d, Len(d))
    ' return value is the new reference count
    My_AddRef = d.RefCount
End Function

' ULONG STDMETHODCALLTYPE Release(THIS)
' Called when decrementing reference count (destroy when to 0)
Private Function My_Release(ByVal This As LongPtr) As Long
    Dim d As MyClassData
    ' copy the instance data into d first,
    ' and restore to d after decrementing reference count
    Call CopyMemory(d, ByVal This, Len(d))
    d.RefCount = d.RefCount - 1
    Call CopyMemory(ByVal This, d, Len(d))
    ' call CoTaskMemFree when the reference count becomes 0
    If d.RefCount = 0 Then
        Call CoTaskMemFree(This)
        ' call exit function
        Call OnExit
    End If
    ' return value is the new reference count
    My_Release = d.RefCount
End Function

' Registers Handler object to call Handler.OnExit when exiting
' Arguments:
'   Handler: The object that implements OnExit() procedure or the function pointer (specify with AddressOf) with no arguments
'   Key: The collection key (can be specified to RemoveExitHandler)
Public Sub AddExitHandler(ByVal Handler As Variant, Optional ByVal Key As String)
    Dim ptr As LongPtr
    Dim vt As VbVarType
    vt = VarType(Handler)
#If Win64 Then
    If vt <> vbObject And vt <> vbLongLong Then
#Else
    If vt <> vbObject And vt <> vbLong Then
#End If
        Call Err.Raise(13)
    End If
    If Not m_collHandlers Is Nothing Then
        On Error Resume Next
        Dim o As Object
        Set o = m_collHandlers.Item(Key)
        If Not o Is Nothing Then
            Call m_collHandlers.Remove(Key)
        Else
            ptr = m_collHandlers.Item(Key)
            If ptr <> 0 Then
                Call m_collHandlers.Remove(Key)
            End If
        End If
        Call Err.Clear
        On Error GoTo 0
    End If
    If m_unk Is Nothing Then
        Dim p As LongPtr
        ' Create a new instance
        p = CreateInstanceMemory()
        If p = 0 Then Exit Sub
        Dim unk As IUnknown
        ' Set unk to the instance of p
        Call CopyMemory(unk, p, Len(p))
        ' Set it to m_unk (implicitly call My_AddRef)
        Set m_unk = unk
        Set m_collHandlers = New Collection
    End If
    Call m_collHandlers.Add(Handler, Key)
End Sub

' Unregisters Handler object
Public Sub RemoveExitHandler(ByVal Handler As Variant)
    If m_collHandlers Is Nothing Then Exit Sub
    If VarType(Handler) = vbString Then
        On Error Resume Next
        Call m_collHandlers.Remove(Handler)
        Exit Sub
    End If
    Dim vt As VbVarType
    vt = VarType(Handler)
#If Win64 Then
    If vt <> vbObject And vt <> vbLongLong Then
#Else
    If vt <> vbObject And vt <> vbLong Then
#End If
        Call Err.Raise(13)
    End If
    Dim v As Variant, i As Long
    On Error Resume Next
    For i = 1 To m_collHandlers.Count
        Set v = m_collHandlers.Item(i)
        If Err.Number = 0 Then
            If v Is Handler And Err.Number = 0 Then
                Call m_collHandlers.Remove(i)
                Exit For
            End If
        Else
            Call Err.Clear
            v = m_collHandlers.Item(i)
            If v = Handler And Err.Number = 0 Then
                Call m_collHandlers.Remove(i)
                Exit For
            End If
        End If
    Next i
End Sub

' Writing exit process when terminating VBA
Private Sub OnExit()
    Dim o As Variant
    On Error Resume Next
    For Each o In m_collHandlers
        Dim vt As VbVarType
        vt = VarType(o)
        If vt = vbObject Then
            Call o.OnExit
        Else
            Dim vt2 As Integer
            Dim v As Variant, v2 As Variant
            Call DispCallFunc(0, CLngPtr(o), 4, vbLong, 0, vt2, VarPtr(v), v2)
        End If
    Next o
End Sub
