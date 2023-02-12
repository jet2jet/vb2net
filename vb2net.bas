Attribute VB_Name = "vb2net"
' Copyright (C) 2023 jet
' For more information about license, see LICENSE.
Option Explicit

Private Declare PtrSafe Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" _
    (ByVal lpModuleName As String) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" _
    (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32.dll" _
    (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function GetProcAddress Lib "kernel32.dll" _
    (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function SearchPath Lib "kernel32.dll" Alias "SearchPathA" _
    (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, _
    ByVal nBufferLength As Long, ByVal lpBuffer As String, ByRef lpFilePart As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
    (ByRef Dest As Any, ByRef Src As Any, ByVal Length As Long)
Private Declare PtrSafe Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageW" _
    (ByVal dwFlags As Long, ByVal lpSource As LongPtr, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
    ByRef lpBuffer As Any, ByVal nSize As Long, ByVal Arguments As LongPtr) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function LocalFree Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" _
    (ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Integer, _
    ByRef prgpvarg As LongPtr, _
    ByRef pvargResult As Variant) As Long
Private Declare PtrSafe Function GetErrorInfo Lib "oleaut32.dll" (ByVal dwReserved As Long, ByRef pperrinfo As IUnknown) As Long

Private Enum hostfxr_delegate_type
    hdt_com_activation = 0
    hdt_load_in_memory_assembly
    hdt_winrt_activation
    hdt_com_register
    hdt_com_unregister
    hdt_load_assembly_and_get_function_pointer
    hdt_get_function_pointer
End Enum

Private m_handleHostFXR As LongPtr
Private m_pfnLoadAssembly As LongPtr
Private m_pfnLoadAssemblyFromFile As LongPtr
Private m_pfnClose As LongPtr

' Call object's method with index of vftable
Private Function VBCallAbsoluteObject(ByVal Object As IUnknown, _
    ByVal IndexForVftable As Integer, _
    ByVal RetType As VbVarType, _
    ParamArray Arguments() As Variant) As Variant
    If Object Is Nothing Then
        Call Err.Raise(5)
    End If
    Dim hr As Long
    Dim argVt() As Integer
    Dim argsPtr() As LongPtr
    Dim i As Long, c As Long
    Dim lb As Long, ub As Long
    lb = LBound(Arguments)
    ub = UBound(Arguments)
    c = ub - lb + 1
    If c > 0 Then
        ReDim argVt(lb To ub)
        ReDim argsPtr(lb To ub)
        For i = lb To ub
            argVt(i) = VarType(Arguments(i))
            argsPtr(i) = VarPtr(Arguments(i))
        Next i
        hr = DispCallFunc(ObjPtr(Object), _
            CLngPtr(IndexForVftable) * Len(argsPtr(0)), _
            4, _
            CInt(RetType), _
            c, _
            argVt(lb), _
            argsPtr(lb), _
            VBCallAbsoluteObject)
    Else
        ReDim argVt(0)
        ReDim argsPtr(0)
        hr = DispCallFunc(ObjPtr(Object), _
            CLngPtr(IndexForVftable) * Len(argsPtr(0)), _
            4, _
            CInt(RetType), _
            0, _
            argVt(0), _
            argsPtr(0), _
            VBCallAbsoluteObject)
    End If
    If hr < 0 Then Call Err.Raise(hr)
End Function

Private Sub LogErrorInfo(ByVal Number As Long)
    Dim hr As Long
    Dim p As IUnknown
    hr = GetErrorInfo(0, p)
    If hr < 0 Or hr = 1 Then
        Debug.Print Hex$(Number); " "; Err.Description
        Exit Sub
    End If
    Dim v As Variant, Source As String, Description As String
    v = VBCallAbsoluteObject(p, 4, vbLong, VarPtr(Source))
    v = VBCallAbsoluteObject(p, 5, vbLong, VarPtr(Description))
    Debug.Print Hex$(Number); " "; Description; "("; Source; ")"
End Sub

Private Function CompareVersionString(ByRef a() As String, ByRef b() As String) As Integer
    If CInt(a(0)) <> CInt(b(0)) Then
        CompareVersionString = CInt(a(0)) - CInt(b(0))
        Exit Function
    End If
    If UBound(a) = 0 Then
        CompareVersionString = IIf(UBound(b) = 0, 0, -1)
        Exit Function
    ElseIf UBound(b) = 0 Then
        CompareVersionString = 1
        Exit Function
    End If
    If CInt(a(1)) <> CInt(b(1)) Then
        CompareVersionString = CInt(a(1)) - CInt(b(1))
        Exit Function
    End If
    If UBound(a) = 1 Then
        CompareVersionString = IIf(UBound(b) = 1, 0, -1)
        Exit Function
    ElseIf UBound(b) = 1 Then
        CompareVersionString = 1
        Exit Function
    End If
    If CInt(a(2)) <> CInt(b(2)) Then
        CompareVersionString = CInt(a(2)) - CInt(b(2))
        Exit Function
    End If
    If UBound(a) = 2 Then
        CompareVersionString = IIf(UBound(b) = 2, 0, -1)
        Exit Function
    ElseIf UBound(b) = 2 Then
        CompareVersionString = 1
        Exit Function
    End If
    CompareVersionString = CInt(a(3)) - CInt(b(3))
End Function

Private Sub RaiseWin32Error(ByVal Error As Long)
    Dim ptr As LongPtr
    Dim s As String, ln As Long
    Call FormatMessage(&H1100&, 0, Error, 0, ptr, 0, 0)
    ln = lstrlenW(ptr)
    s = String$(ln, 0)
    Call CopyMemory(ByVal StrPtr(s), ByVal ptr, 2 * ln)
    Call LocalFree(ptr)
    Call Err.Raise(&H80070000 + Error, , s)
End Sub

Private Function SearchHostFXR() As String
    ' search 'dotnet.exe'
    Dim lnLength As Long
    Dim strBuffer As String, p As LongPtr
    strBuffer = String$(260, 0)
    lnLength = SearchPath(vbNullString, "dotnet.exe", vbNullString, 260, strBuffer, p)
    If lnLength = 0 Then
        SearchHostFXR = ""
        Exit Function
    End If
    ' If in C++ I use 'p', but in VB 'p' is not a valid address, so
    ' I search '\' to extract directory path name
    Dim strDotNetPath As String
    strDotNetPath = Left$(strBuffer, lnLength)
    lnLength = InStrRev(strDotNetPath, "\")
    If lnLength > 0 Then
        ' including '\'
        strDotNetPath = Left$(strDotNetPath, lnLength)
    Else
        strDotNetPath = ""
    End If
    Dim strHostFXRBasePath As String
    Dim sLargestVersion() As String
    Dim strTargetLargest As String
    ReDim sLargestVersion(0)
    sLargestVersion(0) = "0"
    strTargetLargest = ""
    strHostFXRBasePath = strDotNetPath + "host\fxr\"
    strBuffer = Dir(strHostFXRBasePath + "*.*", vbDirectory)
    Do While strBuffer <> ""
        Dim s() As String
        If strBuffer <> "." And strBuffer <> ".." Then
            On Error Resume Next
            s = Split(strBuffer, ".")
            If Err.Number = 0 Then
                If CompareVersionString(sLargestVersion, s) < 0 Then
                    sLargestVersion = s
                    strTargetLargest = strHostFXRBasePath + strBuffer
                End If
            End If
            Call Err.Clear
            On Error GoTo 0
        End If
        strBuffer = Dir()
    Loop
    If strTargetLargest <> "" Then
        strTargetLargest = strTargetLargest + "\hostfxr.dll"
    End If
    SearchHostFXR = strTargetLargest
End Function

Public Sub InitializeVb2net(ByVal vb2netFile As String)
    If m_handleHostFXR <> 0 Then
        Exit Sub
    End If
    Dim strHostFXR As String
    strHostFXR = SearchHostFXR()
    If strHostFXR = "" Then
        Call Err.Raise(53)
        Exit Sub
    End If
    Dim hInstHostFXR As LongPtr
    hInstHostFXR = GetModuleHandle(strHostFXR)
    If hInstHostFXR = 0 Then
        hInstHostFXR = GetModuleHandle("hostfxr.dll")
        If hInstHostFXR <> 0 Then
            Call Err.Raise(5, "hostfxr.dll is already initialized with different version")
            Exit Sub
        End If
        hInstHostFXR = LoadLibrary(strHostFXR)
        If hInstHostFXR = 0 Then
            Call RaiseWin32Error(Err.LastDllError)
            Exit Sub
        End If
    End If
    Dim pfnInitialize As LongPtr
    Dim pfnGetRuntimeDelegate As LongPtr
    pfnInitialize = GetProcAddress(hInstHostFXR, "hostfxr_initialize_for_runtime_config")
    m_pfnClose = GetProcAddress(hInstHostFXR, "hostfxr_close")
    pfnGetRuntimeDelegate = GetProcAddress(hInstHostFXR, "hostfxr_get_runtime_delegate")
    Dim vb2netRuntimeConfig As String
    Dim i As Integer
    i = InStrRev(vb2netFile, ".")
    If i > 0 Then
        vb2netRuntimeConfig = Left$(vb2netFile, i - 1) + ".runtimeconfig.json"
    Else
        vb2netRuntimeConfig = vb2netFile + ".runtimeconfig.json"
    End If
    Dim e As Long
    Dim avt() As Integer, avptr() As LongPtr, avarg() As Variant, vr As Variant
    Dim handle As LongPtr
    ReDim avt(2), avarg(2), avptr(2)
    avt(0) = VarType(handle) ' long-ptr var type
    avarg(0) = StrPtr(vb2netRuntimeConfig)
    avptr(0) = VarPtr(avarg(0))
    avt(1) = VarType(handle)
    avarg(1) = CLngPtr(0)
    avptr(1) = VarPtr(avarg(1))
    avt(2) = VarType(handle)
    avarg(2) = VarPtr(handle)
    avptr(2) = VarPtr(avarg(2))
    ' 1: CC_CDECL
    e = DispCallFunc(0, pfnInitialize, 1, vbLong, 3, avt(0), avptr(0), vr)
    If e < 0 Then
        Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If
    e = vr
    If e < 0 Then
        Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If

    Dim pfnLoadAssemblyAndGetFunctionPointer As LongPtr
    avt(0) = VarType(handle)
    avarg(0) = handle
    avt(1) = vbLong
    avarg(1) = hdt_load_assembly_and_get_function_pointer
    avt(2) = VarType(handle)
    avarg(2) = VarPtr(pfnLoadAssemblyAndGetFunctionPointer)
    ' 1: CC_CDECL
    e = DispCallFunc(0, pfnGetRuntimeDelegate, 1, vbLong, 3, avt(0), avptr(0), vr)
    If e < 0 Then
        Call hostfxr_close(handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If
    e = vr
    If e < 0 Then
        Call hostfxr_close(handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If

    Dim strAssemblyFile As String, strTypeName As String, strMethodName As String
    strAssemblyFile = vb2netFile
    strTypeName = "vb2net.Global, vb2net"
    strMethodName = "LoadAssembly"
    ReDim avt(5), avarg(5), avptr(5)
    avt(0) = VarType(handle) ' long-ptr var type
    avarg(0) = StrPtr(strAssemblyFile) ' assembly_path
    avptr(0) = VarPtr(avarg(0))
    avt(1) = VarType(handle) ' long-ptr var type
    avarg(1) = StrPtr(strTypeName) ' type_name
    avptr(1) = VarPtr(avarg(1))
    avt(2) = VarType(handle)
    avarg(2) = StrPtr(strMethodName) ' method_name
    avptr(2) = VarPtr(avarg(2))
    avt(3) = VarType(handle)
    avarg(3) = CLngPtr(-1) ' delegate_type_name (-1: UNMANAGEDCALLERSONLY_METHOD)
    avptr(3) = VarPtr(avarg(3))
    avt(4) = VarType(handle)
    avarg(4) = CLngPtr(0) ' reserved
    avptr(4) = VarPtr(avarg(4))
    avt(5) = VarType(handle)
    avarg(5) = VarPtr(m_pfnLoadAssembly) ' delegate
    avptr(5) = VarPtr(avarg(5))

    ' 4: CC_STDCALL
    e = DispCallFunc(0, pfnLoadAssemblyAndGetFunctionPointer, 4, vbLong, 6, avt(0), avptr(0), vr)
    If e < 0 Then
        Call hostfxr_close(handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If
    e = vr
    If e < 0 Then
        Call LogErrorInfo(e)
        Call hostfxr_close(handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If

    strMethodName = "LoadAssemblyFromFile"
    avarg(2) = StrPtr(strMethodName) ' method_name
    avarg(5) = VarPtr(m_pfnLoadAssemblyFromFile) ' delegate
    ' 4: CC_STDCALL
    e = DispCallFunc(0, pfnLoadAssemblyAndGetFunctionPointer, 4, vbLong, 6, avt(0), avptr(0), vr)
    If e < 0 Then
        Call hostfxr_close(handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If
    e = vr
    If e < 0 Then
        Call LogErrorInfo(e)
        Call hostfxr_close(handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If

    m_handleHostFXR = handle
    Call AddExitHandler(AddressOf OnExitHostFXR)
End Sub

Private Sub hostfxr_close(ByVal handle As LongPtr)
    Dim avt() As Integer, avptr() As LongPtr, avarg() As Variant, vr As Variant
    Dim e As Long
    ReDim avt(0), avarg(0), avptr(0)
    avt(0) = VarType(handle)
    avarg(0) = handle
    avptr(0) = VarPtr(avarg(0))
    e = DispCallFunc(0, m_pfnClose, 1, vbLong, 1, avt(0), avptr(0), vr)
End Sub

Public Function LoadAssembly(ByVal AssemblyName As String) As Object
    If m_handleHostFXR = 0 Then
        Call Err.Raise(13, , "Must be initialized first")
        Exit Function
    End If
    Dim avt() As Integer, avptr() As LongPtr, avarg() As Variant, vr As Variant
    Dim e As Long
    ReDim avt(1), avarg(1), avptr(1)
    avt(0) = VarType(avptr(0))
    avarg(0) = StrPtr(AssemblyName)
    avptr(0) = VarPtr(avarg(0))
    avt(1) = VarType(avptr(1))
    avarg(1) = VarPtr(LoadAssembly)
    avptr(1) = VarPtr(avarg(1))
    e = DispCallFunc(0, m_pfnLoadAssembly, 4, vbLong, 2, avt(0), avptr(0), vr)
    If e < 0 Then
        Call Err.Raise(e)
        Exit Function
    End If
    e = vr
    If e < 0 Then
        Call Err.Raise(e)
        Exit Function
    End If
End Function

Public Function LoadAssemblyFromFile(ByVal AssemblyPath As String) As Object
    If m_handleHostFXR = 0 Then
        Call Err.Raise(13, , "Must be initialized first")
        Exit Function
    End If
    Dim avt() As Integer, avptr() As LongPtr, avarg() As Variant, vr As Variant
    Dim e As Long
    ReDim avt(1), avarg(1), avptr(1)
    avt(0) = VarType(avptr(0))
    avarg(0) = StrPtr(AssemblyPath)
    avptr(0) = VarPtr(avarg(0))
    avt(1) = VarType(avptr(1))
    avarg(1) = VarPtr(LoadAssemblyFromFile)
    avptr(1) = VarPtr(avarg(1))
    e = DispCallFunc(0, m_pfnLoadAssemblyFromFile, 4, vbLong, 2, avt(0), avptr(0), vr)
    If e < 0 Then
        Call Err.Raise(e)
        Exit Function
    End If
    e = vr
    If e < 0 Then
        Call Err.Raise(e)
        Exit Function
    End If
End Function

Private Sub OnExitHostFXR()
    If m_handleHostFXR = 0 Then
        Exit Sub
    End If
    Call hostfxr_close(m_handleHostFXR)
    m_handleHostFXR = 0
End Sub

Public Sub Sample()
    ' Make vb2net file path
    Dim vb2netFile As String
    Dim i As Integer
    i = InStrRev(ThisWorkbook.FullName, "\")
    If i > 0 Then
        vb2netFile = Left$(ThisWorkbook.FullName, i) + "vb2net.dll"
    Else
        vb2netFile = "vb2net.dll"
    End If

    ' Initializes vb2net
    Call InitializeVb2net(vb2netFile)

    ' Get type of System.String
    Dim asmMscorlib As Object, typeString As Object
    Set asmMscorlib = LoadAssembly("mscorlib")
    Set typeString = asmMscorlib.GetType_2("System.String")
    ' Get constructor of System.Text.RegularExpressions.Regex (with parameters: System.String)
    Dim asmRegex As Object, typeRegex As Object, ctorRegex As Object
    Set asmRegex = LoadAssembly("System.Text.RegularExpressions")
    Set typeRegex = asmRegex.GetType_2("System.Text.RegularExpressions.Regex")
    Set ctorRegex = typeRegex.GetConstructor(Array(typeString))
    ' Create System.Text.RegularExpressions.Regex instance
    Dim regex As Object
    Set regex = ctorRegex.Invoke_3(Array("(\d+) (\d+)"))
    ' Executes Match
    Dim m As Object
    Set m = regex.Match_4("123 456")
    Dim v As Variant
    Debug.Print m.Groups.Count
    For Each v In m.Groups
        Debug.Print v.Value
    Next v
End Sub
