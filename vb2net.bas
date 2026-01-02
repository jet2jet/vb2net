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

Private Declare PtrSafe Function CryptStringToBinary Lib "Crypt32.dll" Alias "CryptStringToBinaryW" ( _
    ByVal pszString As LongPtr, _
    ByVal cchString As Long, _
    ByVal dwFlags As Long, _
    ByVal pbBinary As LongPtr, _
    ByVal pcbBinary As LongPtr, _
    ByVal pdwSkip As LongPtr, _
    ByVal pdwFlags As LongPtr _
    ) As Long
Private Const CRYPT_STRING_BASE64 As Long = &H1&

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As LongPtr
  bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

Private Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Declare PtrSafe Function CreatePipe Lib "kernel32.dll" _
    (ByRef hReadPipe As LongPtr, ByRef hWritePipe As LongPtr, ByRef lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare PtrSafe Function SetHandleInformation Lib "kernel32" _
    (ByVal hObject As LongPtr, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function ReadFile Lib "kernel32.dll" _
    (ByVal hFile As LongPtr, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare PtrSafe Function WriteFile Lib "kernel32.dll" _
    (ByVal hFile As LongPtr, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As Any) As Long
Private Declare PtrSafe Function PeekNamedPipe Lib "kernel32" _
    (ByVal hNamedPipe As LongPtr, _
    ByRef lpBuffer As Any, _
    ByVal nBufferSize As Long, _
    ByRef lpBytesRead As Long, _
    ByRef lpTotalBytesAvail As Long, _
    ByRef lpBytesLeftThisMessage As Long) As Long
Private Declare PtrSafe Function CreateProcess Lib "kernel32.dll" Alias "CreateProcessA" _
    (ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, _
    ByRef lpProcessAttributes As Any, _
    ByRef lpThreadAttributes As Any, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByRef lpEnvironment As Any, _
    ByVal lpCurrentDirectory As String, _
    ByRef lpStartupInfo As STARTUPINFO, _
    ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32.dll" (ByVal Handle As LongPtr) As Long
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32.dll" _
    (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32.dll" _
    (ByVal hProcess As LongPtr, ByRef lpExitCode As Long) As Long
Private Declare PtrSafe Function TerminateProcess Lib "kernel32.dll" _
    (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long

Private Const HANDLE_FLAG_INHERIT = &H1
Private Const STARTF_USESHOWWINDOW As Long = &H1&
Private Const STARTF_USESTDHANDLES As Long = &H100&
Private Const CREATE_NO_WINDOW As Long = &H8000000
Private Const WAIT_OBJECT_0 As Long = 0

Private Enum hostfxr_delegate_type
    hdt_com_activation = 0
    hdt_load_in_memory_assembly
    hdt_winrt_activation
    hdt_com_register
    hdt_com_unregister
    hdt_load_assembly_and_get_function_pointer
    hdt_get_function_pointer
    ' from .NET 8
    hdt_load_assembly
    hdt_load_assembly_bytes
End Enum

Private m_handleHostFXR As LongPtr
Private m_pfnLoadAssembly As LongPtr
Private m_pfnLoadAssemblyFromFile As LongPtr
Private m_pfnClose As LongPtr

Private Const BIN_VB2NET As String = ""

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

Private Function Base64StringToByte(ByRef sData As String) As Byte()
    Base64StringToByte = ""
    If Len(sData) = 0 Then
        Exit Function
    End If
    
    Dim pszString As LongPtr
    Dim cchString As Long
    pszString = StrPtr(sData)
    cchString = Len(sData)
    
    Dim nBufferSize As Long
    Dim bBuffer() As Byte
    
    If CryptStringToBinary(pszString, cchString, CRYPT_STRING_BASE64, 0, VarPtr(nBufferSize), 0, 0) Then
        If nBufferSize Then
            ReDim bBuffer(0 To nBufferSize - 1)
            If CryptStringToBinary(pszString, cchString, CRYPT_STRING_BASE64, VarPtr(bBuffer(0)), VarPtr(nBufferSize), 0, 0) Then
                 Base64StringToByte = bBuffer
            End If
        End If
    End If
End Function

Private Function ExecuteWithBinaryPipe(ByVal CommandLine As String, ByRef StandardInput() As Byte) As Byte()
    Dim hpReadStdIn As LongPtr, hpWriteStdIn As LongPtr
    Dim hpReadStdOut As LongPtr, hpWriteStdOut As LongPtr
    Dim sa As SECURITY_ATTRIBUTES
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim ln As Long

    Dim dw As Long
    dw = UBound(StandardInput) - LBound(StandardInput) + 1

    sa.nLength = Len(sa)
    sa.bInheritHandle = 1
    ln = CreatePipe(hpReadStdIn, hpWriteStdIn, sa, dw)
    If ln = 0 Then
        Exit Function
    End If
    ' This is necessary to work pipe correctly
    Call SetHandleInformation(hpWriteStdIn, HANDLE_FLAG_INHERIT, 0)
    ln = CreatePipe(hpReadStdOut, hpWriteStdOut, sa, 0)
    If ln = 0 Then
        Call CloseHandle(hpReadStdIn)
        Call CloseHandle(hpWriteStdIn)
        Exit Function
    End If
    Call SetHandleInformation(hpReadStdOut, HANDLE_FLAG_INHERIT, 0)

    si.cb = Len(si)
    si.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
    si.hStdInput = hpReadStdIn
    si.hStdOutput = hpWriteStdOut
    si.wShowWindow = 0 ' SW_HIDE

    ln = CreateProcess(vbNullString, CommandLine, ByVal vbNullString, ByVal vbNullString, True, _
        CREATE_NO_WINDOW, ByVal vbNullString, vbNullString, si, pi)
    If ln = 0 Then
        Call CloseHandle(hpReadStdIn)
        Call CloseHandle(hpReadStdOut)
        Call CloseHandle(hpWriteStdIn)
        Call CloseHandle(hpWriteStdOut)
        Exit Function
    End If

    Call CloseHandle(pi.hThread)

    Call WriteFile(hpWriteStdIn, StandardInput(LBound(StandardInput)), dw, dw, ByVal vbNullString)
    Call CloseHandle(hpWriteStdIn)

    Dim by() As Byte
    Dim curBufferLen As Long, Offset As Long
    ReDim by(0 To 8191)
    curBufferLen = 8192
    Offset = 0

    Dim r As Long
    Do
        DoEvents
        r = WaitForSingleObject(pi.hProcess, 10)
        If r = WAIT_OBJECT_0 Then
            Exit Do
        End If

        dw = 0

        If PeekNamedPipe(hpReadStdOut, 0&, 0, 0&, dw, ByVal 0) = 0 Then
            Call TerminateProcess(pi.hProcess, &HFFFFFFFF)
            Exit Do
        End If
        
        If dw > 0 Then
            ln = ReadFile(hpReadStdOut, by(Offset), 8192 - (Offset Mod 8192), dw, ByVal 0)
            If ln = 0 Then
                Offset = Offset + dw
                Call TerminateProcess(pi.hProcess, &HFFFFFFFF)
                Exit Do
            End If
            Offset = Offset + dw
            If Offset = curBufferLen Then
                curBufferLen = curBufferLen + 8192
                ReDim Preserve by(0 To (curBufferLen - 1))
            End If
        End If
    Loop
    Call GetExitCodeProcess(pi.hProcess, r)
    Call CloseHandle(pi.hProcess)
    Call CloseHandle(hpReadStdIn)
    Call CloseHandle(hpWriteStdOut)
    If r <> 0 Then
        Call CloseHandle(hpReadStdOut)
        Exit Function
    End If

    Do
        dw = 0
        ln = ReadFile(hpReadStdOut, by(Offset), 8192 - (Offset Mod 8192), dw, ByVal 0)
        If ln = 0 Then
            Offset = Offset + dw
            Exit Do
        End If
        Offset = Offset + dw
        If Offset = curBufferLen Then
            curBufferLen = curBufferLen + 8192
            ReDim Preserve by(0 To (curBufferLen - 1))
        End If
    Loop
    ReDim Preserve by(0 To (Offset - 1))
    Call CloseHandle(hpReadStdOut)
    ExecuteWithBinaryPipe = by
End Function

Private Function GetVb2netBinary() As Byte()
    Dim bin As String
    bin = ""
    ' The following base64-encoded data is generated by 'vb2net\makeBinCode.bat'
    bin = bin + "H4sIAO1jV2kAA+1aDXBc1XU+9723b9/uSmu9XXklY9le4x/W+rMkyzI2xliWZFvE+rH+bIOJvJKe5cWrfcvblY1iTOQ6pCXEBE9SaCBQJiWZhGZa0oQJKQnNDE0JhTROSGjIEDc0aQrjpskkTSEZAv3OfW9XWtkl6dBppx2etOeec+65556fe8+7b/eNpTLrB6y0lcxZ6zNWvq2xqeF4KjNhH8+tPzbWAk7jRDpNb+9qampqa2uLc4trYbuhmYnm"
    bin = bin + "1g2bWpqbW5tamuJNzRs3tW6geNPbnPd3uqZz+aQDU96unoXO/R+5eq67i1S0Gj5vvkn0mMff/juMncUnvOIvw/T5wNdXPib2fH3l0JFULp517EknORUfT2Yydj4+ZsWd6Uw8lYl39g3Gp+wJq7G8PLja09HfRbRHqLTuqUe+V9D7A7o8HhLISC0I3eX9+24AjuohSVZIXHHtJppr6UGXz5dKh25jUf6fa4uNvB6C3j5y9SbUSztZhuZru4iGfoeY"
    bin = bin + "FC/YZ8wjDdC759GNeevmPNrXE55ftXN2z1NxqNHJOePk2XbIc7S+VA652t7oWGl73LVVJoZ1NV0kt2Ohmd/1jNoth/joI3VErTVE4r/i67xruZKIEgVr3+22StUZOCjWxNSPnYRKrXaLx1cXkpzLWoyu5NFE0SaVPkPSDtNlQqYcMkpuMYiTYaCqHQOaqwII2tWAoVhZ3XG/cW95wF4C0r4MIKrpphbx2UuB119h+qr2R33gmD4bXgZNzcEM2Ygv"
    bin = bin + "sQxUYrkcUBeAQJXsV07wRBH9jcV+xCcC/smQZAROItKaqdsrIGVU7S8z/HemWk7XBmO1l+WQuOB72AeNWoRMBeZdr8dVUmZrtdg9iZUQWNOgra2s8yXYyGAsdOUvOQNyOr99OXhlxgmeSk+sArH5OfSGfTF/1f4HjBM89TQcElHN1CoTq9k1uGKvAVJed4UnwFmM6qZemVjrCuj2FVJAT8CouvJY1FcV9dcd1Uz/xRZFjVg0EIsG61oCZvBee92C"
    bin = bin + "btOQZtWHwz7TbwZhVVU0UFfOWDRoBgOIxVNmoCFQFfWxxOIHTFge9Zt+X8tdpq9ej+scilgtuCG/67WNrAev+fUbb77pxoloaZOfVqq8j8lUV7y+oUs5wStgegU75jN9DV5Ekcs6XghYAnoCeyMY1hOIeV3VmUa59laU3x8L61UJUEiPWtUSVudN2expcVlq1QEZ/TVzQQ7OhTMU9vFGrCurLNe9uO/T3bCHtarRgkFwnwNfdQDhrQqY/nvtJhZ1"
    bin = bin + "WTpCcYBDEVv/I+NEMwdTbwm5GAY23xwPsw6/l7aU381aWFsRLpnAAKdsFEqx+qKGaUQD4EJzkKcMFqcMgAWaO8zg3JSBlqCHGQiAnDHQoMbqVCjJtchtENhcjnXiSsW2CeCx6KK6MFacvYEHLsK8i8xFnOsLxolWsM7UcmZCLlFfaZzYyHKhmMtItPFWkSiswfLeJGlWn7iSZyxfHA27NPor6g6ZFVHTLIu1NJshs+wj0agZrWpZZEYxYZMB/EDU"
    bin = bin + "rKswoyt+s0FnyTrNLIuGzTLWbZabFfcGTPNeezNrX4SFEY1g05tmxN7CNlXwDGZFbP19ZjjWkggbt17Fglul3VXSjzUxM/xALGxUJa4GN67JNVnI+uW6uxax6bZxpblGrjrd3s5acH/R6w4WEjjo5qneXxgRs9tl5anC4B3SWbuDdUhFnQB1cX+p9i6p3W/vLGiPRSvPh8L+RIQdqzxPZiXfPoS8sf2+8sGoGopzJlBDn3BvGabSoFTWKrJgJkyM"
    bin = bin + "qq9WKxO72L5yRU3s5hUWqFcCtUqtMg35pUG9/jDXMZ1LjNqgVtbpqt3NlTfORcnU3tBRMvTctawQA+qwIbUT3OepXeORpu/OVGIPp1c3fSa2gN3DQ3rlusfWxXwoc0rIX+/3n1jBNwFlGjdOpcyo9xsyCrWVtW4tvdO7x2iuBwYM7+O7hdKwtNJFizs+cH9MdTd8rXZuPXiJfp51L0BDRDu3oYSDEHwCq1KoCR8cqVVyQag/h4OH4soMAKilFWWe"
    bin = bin + "frZNodXeDVM5t7FE9xbeQ8q5tlLma1JyUynzh5LZWMr8O8msLWV+UTJbS5kPSWailPkhybyyNAAnlXObSzlHlXNbSjkHlHNXlXJ2Kee2lnI2KueuLuWsUs6tK+WYyrltJZzaqlpel6dgGKJt1hmKPQh2Bbltw8+UxBDfEdV5O/JARLyxGPtORIRiD/MiHAFgYp+8L+lrHejKntzvcQ+UcK+T3NjJ62WrnzyI1q/GEkhwnV47ci5YTPS7AWaXe+cS"
    bin = bin + "78whlyTfizT6GsnjqKnKAhmsW67IBWuPso2X6R51iHcsF0h1TdWBChFrekG3kyBDdVdGRGJM3tk1o6HsXLQQlzrdcDcmb20tgVNc8Fa24sb9iQkOjYJxFsfmr86HEYkYbFhqHwbjRd+8IFUVypYuq8oDMZ9buVyyQasqP18WKz9vyPCWn6dArajhs8km6rqFwoXz2xNPyKOqGKsrm+fd2oIvF2pbPPYk5zLjEvLOey5QsmG8mGWg08/1x7X0CJtS"
    bin = bin + "rs4ROAqo2wmVVZFHkKDeYEhG2fnNeiIFzqU2Y6jqzHrpqv8+3oY38g2qzPOs7DwZtYVauJI6Jlx/3rYtR//7bCmccgttG44jXCrlfQryOwav3SG8Ezif54+1NjY1bsCD8maSu4afx7ejaK66legzaB9ifDDvpDKTOZbor3KfPVYND9L+Kvd5Z9Wu4W7cW+gI6GdRr1ftSNtjXs5R2cS+a5RgAO7Rr8UGirnnf34GRAGWjycICFaK+1iBUBAqMPGj"
    bin = bin + "S8iNqZxPd8/r8nlSUOHR7Qnd9USnjb5HDZ1ekLBZO20sou8YzL9Na9F1+orvVcCYhBkJn9MY/kzCYcnZJ/HHpfwDEv5MZW0/8b0M/HyA4RKNYSr4sh6ka9S7jSBlNIYhQJ0aFO69XmX4CT9re0TCGyR8AjBIWVEZCOOB7CWh0weUewI6bfW/rIfx/Mecv5BjUyHW3w3JE7ReehE37kEAP+7jWV6T9vTqDHcQw4Cf5T+ic+8zKs/1Psm/xWBYH2K4"
    bin = bin + "W3I+LS08IfknJeenUs+/+Rm+KjWflvx2V4OEX5ejPi7hx6R8ubTzbsnJhV6VyfiAzIRcW1iU59V8oL1IdWmPGkwpkrrdx5QiHzcr6KfY2z3Iq48exGI6qGeDnHmdDEg+L7LBJKiglHwafT1YNUEpeadgyWqKSsnzQZaspiopqRssuQK0AHVGjovTMlIhySuGqTitpa2BLyprqUH/sjIy+wuqECOzg4Lh5yX+xxKS5PytxG+Q+H2AsE/9OeAz/q9A"
    bin = bin + "w1Llr5WV9GjgKWXvbKX4BuA2CddDcu/sIxLGJfwwMTwh8Scl/D6xttbQNxSsXGKYlPAmCbfSEvSmAWP0aLBC8FzfAt9S/17pl98Evb96C/IuaH+R+kdFULZI/TOo95b03Vkyzr3+xrc09Ng8asb/E8WlZkG9bvxC0YrUaYXUAP1RUUtYNamfNzHWcBBaIjToUUegJUIHPeoCtFTSxy53KQdaltCLa5g6i0PhSbGSPrWWqbvpyWBUXUWdV7h9v/HV"
    bin = bin + "qGspn3Cp2/yr1CuKtgyF6tR1Ratf8QtqKFJ3BHVqLFKb1Ra1kYQUPlu93WgDVbWu4MMWUIl1rPMlsQk611N1rWvnd6CzhZo96jh0bqBtHrUGOjfQwVrXajvUDuplj/q+0q5upHzdXKw30S0l1PskdRprdpe6yS2FJOKCRvwirtAqhXdPQmE8GBB4OPuRIfCosEpR4j56XXDvr/y8p1hGlfIaXRsQ2El3AOr0cFCgft4LGYN2Gbw3l/IGhH6utask"
    bin = bin + "zkdP1s+V9UcGn+eZz/pJ6iepn3tXyW/Erg2Q1E9SP0n9VNQf5/sXvRqcw58L8liX8wWV4WrJx6iiv/cH2Yt/hZ0qhRWXM2ct49xLsrfAEXSQT/L0cYn3z8Of8xvY9yFaTQHaYgjcTTjaSwCDtA6wAruM4WYJ2yXslnCvhAckTEqYkvAmCWeknrOAi+nTgEvps64ewfxTsvejEpZ5+IfUrVQpcVM8HLiWnqY3lX7U9WqjH/yQPgT4D/oNtFfs0Zlf"
    bin = bin + "GbQwNqik6YDUmRQ36beB87z+AfSeCt1FL9CZ0N2UEl8OPEAvAX8I8BngS8R9oc/QN6V+3NWxB39F9cYPoKE9+GPAkeAFwFtCrwJ+SSXxTXpILxOv0N5ABQ5qHKUaOe8rtA+cV2gssESwfA3wYZ3x1wJrMOokPKqBzfViRpzzLwE8E9ogJTdDz2Roh0iJ08pu8QJlpCRbXiPtqaGU0ideoir9OsCwvhvwk5AxoSEnTol+9aQwxT3g1MOeu8QS8Tn9"
    bin = bin + "bsAW5c8AD+uPwcdbQo8J9uVJSD6uPA14CmOTIhF8Xgix0v9PwFv9F4DfgDisFDv9PwfeLvE9CuOzOuM/B75SHAhVAO5TKjAqbfwScK2uKc2iP1QmmsWF4FKlBjFcobTShWCTMiPOBjdKuAUclvmVeF5FxRZcZWZgyQYhFJ/SpXA09ihC+TZwA9l4N2AFjQNW0hGF1+YUYA05gHE6DriabgFM0ClZ798PuJ1uB6eTPgi4m84C7sHd1qB+ug9wiB4E"
    bin = bin + "3E+fADxIDytrUHFfVBqx5vegikVpGHAp5QFXIWONVEePA26gpwCvom8DdtAFwHfRa4CDsvd6CcdJ1RpxBlO1Dsphrg7ca+9TxiU+7uEnaIV2WnJOS85pcHZqT9LlOB09iWo2BXgHHdN+KDk/lLapgnFVML5KsMwqcQc9qP2JYD0/QG9UI5VlHlQ/DM5XJHxJjYAjNB4V1y4HnAAe1VT6c8VHn1X8+IToc0o52kX4mMCr0A7SF5RhtNfjs5Ic+hKN"
    bin = bin + "ifMipFyDeMWQEc5BGPAy5IUzEUHkWnAaCFErshCiNpzPQnQlVuJGnLCi4jbxB2Kdos0SlX53zecPfR49rMRUooW8u42LefcEuC3H7Hvpc4j9L2iU/OIe+rS4nr4lHqc33Gm2bhsfHe1M5bLp5ExHOpnLNbeNNtHW4cxxJ5ntG7vRGs9vGxsFa08qlz/UTMcHrJzlHLMmmulYMt1Mw92Z/IYW6rAzxywnbzmHWqgzNZ5P2ZmkMwOiKN/C8vPoDXJo"
    bin = bin + "W6tsmttoa489MZ22tlF370jfu7pGdw73dlBn92D/aNdof/tAe09v39DOvuHezgKzp6tnR9fAQu5w77t6+/b19rb3dBU09Q/09XcNDB3oHx4a6NpJXaPgt+/p7mwf2FUYhc6unu7Bnvahjt00MjTKbIkz0t052tu1r6t3uKdAepOwZAHt2t/R1d/du7MP+vv7unuHugYKynv72BqeblCypDeDC63b1TVE3V2Z6amR9oHu9t6hwugd7Z1yQAe8HCoY"
    bin = bin + "MM+lS3hJPdbUmOV0JvNJGpzJ5a2pxg47nbZkWnKNu6yM5aTGyZnkvHdP0B47OUHtExOU7bDS+Z1WfvyINUG8ALJo0+OpCXJSALumATocK5m3ujO5fDIzbtGklR8dtKcdoGO5vOOhuXF7wu3rsXK55KRFu6z8QDIDxPtZtpvnTmKmxu7MMfuo5fpuOcmxNHDutHMS77SnuRmYzuRTU9bQTNbancxMpKVGpnY69pTHGcQTJBp2Z2fKQ9pzOQQjPcNi"
    bin = bin + "kslW7rbSWUmwhb3JKYsyDBAwawiTFILmzVk0eV52GjvSdsbCsp/C6rWcw8lxaRp1HGEnJTqSTE+7mJW2piTC0w1Y+WknUyT7kw6mhgrJ8XzitguDrEx+TtCxs9hhM5KRZ9Cfd4ZsPDdPj0OjhaTnj9gTO5I5OCLDZ1EXlsHEhDXRnsfT9dh0XhqcheOOXANwd14XZ3eOKmLDnL45fqc1Nj05yZmZ4/VOp9OlnELUh1L5+eySWM2xh5IOPNzJgThu"
    bin = bin + "O0fnOgans1kbVWWib7A/ncwftp2phaMu5hczDj9HLCeHNX9xZ3eGByVlnUr/p1KoaodTk9OOlFsYRMkcsNLJmyWWu3g4UjaB5FxKb3bGSU0eyV8cQ0zJP63OdQxnppIZpGCiI4kt7OT6MumZS2mcyiYz8zq8tSv5+dRYKp3Kz+sd3DHD6cj1ou+YJVcqDTkzWHYuPlhAjrlk6j1W32HamU5OLtgbjV7ssPNILkZGnMljyA0Vt7hcv3uszGT+SHHR"
    bin = bin + "Qm44k5I9vBv3pDJHaVfaHkumqSfp5I6g7bTGU1NoWaY715d1k1W41bSn0zTvFY/CXN1TWW8ErKRcAXGtph22nbaSmaIT1mGvKlKvneehcs9ZE103j1tyPldVBptljoUo7R7Yac/jcE3ptHLjTsqlU0O8vrwdKdGLqp6302VnUdGQLVHJzM6hbjmXaLFcuJRXEySRYnCd5diXrFeDR1NZJGYk6XCNcmhCln5XM3mlN287+2TRd5DLPfZxtN1zXbIo"
    bin = bin + "zVGN4y6UTXa63Znschw6jE3HBRVppqS3Ngt0sVQN2UzyfV+y3Vx0ppKTGTuXT43nFq4xWTbs7CDODalx66LuQk0r9rs1CisMRwqQ42xA7pIp6O7M9R12u53J3Htc7K1n5+k4c7niluVtkaPjbjOOQHhTTkjUmeRQpybmOD2pccfO2YfzUDVhtWNNz+RSOcqyYTK/uUvctTn6brZyNOW1YBXXg6S6MxPWzfNYxQ3v1soc7cbpazqdl8Lu4S6HhVyo"
    bin = bin + "Ak4qmUEfNth4Mk/zj4AsP/9eh0FzHd3DmaMZ+3imhFmic67HTcIlVygfDPNkZ0e7bs6mU+OpPI3jNEJZ12TKQpeHcjmyD3u6qW86n53OD9kF2p7OF9DO6ampmQLBO7lj2nGwweX9c3eqwLSR6bfcoa6EV3jo2KXt7+U3YnrsY5ZEJo5zWfOqObZTu+MkZ0pOJMR3AenvTdNJLs+Md2esAtU1lQXkSx2VDwW5R6Kti317d37yiq9GH39mpI20uBCG"
    bin = bin + "GifhA2KaTIYl2iXh3qjfvzgye5sSNpZFLCNihQ2zRl8cmRFmjS9OkdnbC92gIAFNYSOuKFGqiAp/XBFRPJjopIZxscIoRfwh1rEskopMKZGbjLBiaCTC1eiNkuEjRQmH/YCRqciU1JmS1E2Rm/QKQcuWkxYUkRTmMVZh6qnITeY0ZNiimshJg69l+K9hBEDa+FFMUF3jrxCKWE7ValBUG6SJZbpRbUAn86DTnIZORdENn9Sl6FKoGsoMnZRlNZCF"
    bin = bin + "O0aU2MRqV4MBDWxjlNBXU7OsJhhXeDQcMfxxVTCrwu9fVhOpj8RrIq1AZs8E2J3Z2zl0hEFooZ2j+FVQPsxRUx2pUJRlywWclezltFxoQfRiKBRghBEkPKUyeoZDhiEGgb7dnH0a7gWhuJoV637FMGfPyiCc0kkYaBQ9bFQIlb1WfHB79qzbcfsivyY9X2zO/iEs50VRE4iTJOGgy9bieHBHOIPVHNOgz6+Gq/nr2cgpEZYhqUZIjEeuGX2v+Z3g"
    bin = bin + "Fm0f+Wc1/v1I43c2NH6BQPMx4HdfND8Dfu9CgWPlpL2JS/sNg9cZ8MOlxj+2aPydrsbPp6IGi7W2XxF+VV/mU3SfqkcsRa/xe0tS0XVN54QD3O6TT6e8on289JYBKlgSZXGfUWHeAEuxUsrjyC8oQ6/GxbI1BkKLFCB6JBQ3Sxy3aj2uIt/VCEs1f0GoQxYzwRtDDkTEhPe623L+znNIifGNsNfOzN2ajzj28ZwwhPeW29WCqht7u4Y6bMdqz2br"
    bin = bin + "vXPQ1cfaGpugIFxZPM56D9dc/w0eEYdESJDulUIqE+T3XmEligkyi6fDeKJjXfzGgkxTY3NjUyO/FiPIJwmiRYKC+9z3XTcx7ZPm+eRvImsFve8of2c5pJTjsGQf70HtTOGIA0agO3MEz4A46bDYnsrfLmYUftFbJ2h1U1NTS1NrU1MDvyvqgo4C5l6tbSWSrW8t6V7Xv3j2Aomefg5wEp+a3fjUzn2nwfNH8eFv5k+i7+S8Pr4qSkkaGOwczI3f"
    bin = bin + "/82RxPc6PnXrtfn2f7nqYdaxc8vBISuXj/e7L3oedBNxcCqZyhRwe+zGg15ODpa+VuxJNGYnxmhwd3vLxjby5vhwYQ74cfRMz6nfW/fr8EO3pP/0Q8/Oe3Hyx4X3QC9xfXf+C5aE25LTmU73wCqawvO0Y1nFd5nfXAMdC91955K/xPHvdfIN0hI+573pEny+OOz7DxE9O+/91WeVVsARGqRRwC4aANZNfdQLuhtwp/vWLX1Z++kbBf1zGomu8Vqu"
    bin = bin + "fgtei6VOKTeCNe5AT4rShEcxytBhsmX/ajlqCL1JcHPoT1IecrZ8A4CvR7R6+SvFIPgOevCYcglNhpRpKv610ph8pbaR/OB3QGaKsnKOGXiUBGVJ7ccg1wKuRXlJt0n5gv5OfHDAlfNmS+xaOK4JJXVu3Ag+DkbOyTfBkmZ8muSH5wlBnu3PS9kMbErPs2y+/kaaQF/ay18E4/aAPylHsGdZ+MQWTtIRac3FvDglwF2H9kbP4m1y/j5PJuXNX7A/"
    bin = bin + "81vtcOPaj/E2uNOIUf4t49pE+kXyC6M0FyOiK2U829Gbg9QUtKXhUfwtxjQufIX6/9+13X3H5LUr/7cNeed653rneud65/qfvP4DMqGIaQA4AAA="

    Dim targzBin() As Byte
    targzBin = Base64StringToByte(bin)
    ' Use installed 'tar.exe' to extract DLL binary
    GetVb2netBinary = ExecuteWithBinaryPipe("tar.exe xzO", targzBin)
End Function

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

Private Function MakeTempRuntimeConfig() As String
    Dim FileName As String
    Dim o As Object
    Set o = CreateObject("Scripting.FileSystemObject")
    FileName = o.GetAbsolutePathName(o.GetSpecialFolder(2) + "\" + o.GetBaseName(o.GetTempName()) + ".json")
    Dim FileNum As Integer
    FileNum = FreeFile()
    Open FileName For Output As #FileNum
    ' NOTE: we need .NET 8.0 or later to use 'hdt_load_assembly_bytes'
    Print #FileNum, "{"
    Print #FileNum, "  ""runtimeOptions"": {"
    Print #FileNum, "    ""tfm"": ""net8.0"","
    Print #FileNum, "    ""framework"": {"
    Print #FileNum, "      ""name"": ""Microsoft.NETCore.App"","
    Print #FileNum, "      ""version"": ""8.0.0"""
    Print #FileNum, "    }"
    Print #FileNum, "  }"
    Print #FileNum, "}"
    Close #FileNum
    MakeTempRuntimeConfig = FileName
End Function

Public Sub InitializeVb2net()
    If m_handleHostFXR <> 0 Then
        Exit Sub
    End If

    Dim bin() As Byte
    bin = GetVb2netBinary()
    On Error Resume Next
    If LBound(bin) = UBound(bin) Then
        On Error GoTo 0
        Call Err.Raise(53)
        Exit Sub
    End If
    If Err.Number <> 0 Then
        Call Err.Clear
        On Error GoTo 0
        Call Err.Raise(53)
        Exit Sub
    End If
    On Error GoTo 0

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
    ' Make rutimeconfig.json into the temporary directory
    vb2netRuntimeConfig = MakeTempRuntimeConfig()
    Dim e As Long
    Dim avt() As Integer, avptr() As LongPtr, avarg() As Variant, vr As Variant
    Dim Handle As LongPtr
    ReDim avt(2), avarg(2), avptr(2)
    avt(0) = VarType(Handle) ' long-ptr var type
    avarg(0) = StrPtr(vb2netRuntimeConfig)
    avptr(0) = VarPtr(avarg(0))
    avt(1) = VarType(Handle)
    avarg(1) = CLngPtr(0)
    avptr(1) = VarPtr(avarg(1))
    avt(2) = VarType(Handle)
    avarg(2) = VarPtr(Handle)
    avptr(2) = VarPtr(avarg(2))
    ' 1: CC_CDECL
    e = DispCallFunc(0, pfnInitialize, 1, vbLong, 3, avt(0), avptr(0), vr)
    Call Kill(vb2netRuntimeConfig)
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

    Dim pfnLoadAssemblyBinary As LongPtr
    avt(0) = VarType(Handle)
    avarg(0) = Handle
    avt(1) = vbLong
    avarg(1) = hdt_load_assembly_bytes
    avt(2) = VarType(Handle)
    avarg(2) = VarPtr(pfnLoadAssemblyBinary)
    ' 1: CC_CDECL
    e = DispCallFunc(0, pfnGetRuntimeDelegate, 1, vbLong, 3, avt(0), avptr(0), vr)
    If e < 0 Then
        Call hostfxr_close(Handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If
    e = vr
    If e < 0 Then
        Call hostfxr_close(Handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If

    Dim pfnGetFunctionPointer As LongPtr
    avt(0) = VarType(Handle)
    avarg(0) = Handle
    avt(1) = vbLong
    avarg(1) = hdt_get_function_pointer
    avt(2) = VarType(Handle)
    avarg(2) = VarPtr(pfnGetFunctionPointer)
    ' 1: CC_CDECL
    e = DispCallFunc(0, pfnGetRuntimeDelegate, 1, vbLong, 3, avt(0), avptr(0), vr)
    If e < 0 Then
        Call hostfxr_close(Handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If
    e = vr
    If e < 0 Then
        Call hostfxr_close(Handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If

    ' Load vb2net from binary
    ReDim avt(5), avarg(5), avptr(5)
    avt(0) = VarType(Handle) ' long-ptr var type
    avarg(0) = VarPtr(bin(LBound(bin))) ' assembly_bytes
    avptr(0) = VarPtr(avarg(0))
    avt(1) = VarType(Handle) ' long-ptr var type (size_t)
    avarg(1) = UBound(bin) - LBound(bin) + 1 ' assembly_bytes_len
    avptr(1) = VarPtr(avarg(1))
    avt(2) = VarType(Handle)
    avarg(2) = CLngPtr(0) ' symbols_bytes
    avptr(2) = VarPtr(avarg(2))
    avt(3) = VarType(Handle)
    avarg(3) = CLngPtr(0) ' symbols_bytes_len
    avptr(3) = VarPtr(avarg(3))
    avt(4) = VarType(Handle)
    avarg(4) = CLngPtr(0) ' load_context
    avptr(4) = VarPtr(avarg(4))
    avt(5) = VarType(Handle)
    avarg(5) = CLngPtr(0) ' reserved
    avptr(5) = VarPtr(avarg(5))
    ' 4: CC_STDCALL
    e = DispCallFunc(0, pfnLoadAssemblyBinary, 4, vbLong, 6, avt(0), avptr(0), vr)
    If e < 0 Then
        Call hostfxr_close(Handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If
    e = vr
    If e < 0 Then
        Call LogErrorInfo(e)
        Call hostfxr_close(Handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If

    Dim strTypeName As String, strMethodName As String
    strTypeName = "vb2net.Global, vb2net"
    strMethodName = "LoadAssembly"
    'ReDim avt(5), avarg(5), avptr(5)
    avt(0) = VarType(Handle) ' long-ptr var type
    avarg(0) = StrPtr(strTypeName) ' type_name
    avptr(0) = VarPtr(avarg(0))
    avt(1) = VarType(Handle)
    avarg(1) = StrPtr(strMethodName) ' method_name
    avptr(1) = VarPtr(avarg(1))
    avt(2) = VarType(Handle)
    avarg(2) = CLngPtr(-1) ' delegate_type_name (-1: UNMANAGEDCALLERSONLY_METHOD)
    avptr(2) = VarPtr(avarg(2))
    avt(3) = VarType(Handle)
    avarg(3) = CLngPtr(0) ' load_context
    avptr(3) = VarPtr(avarg(3))
    avt(4) = VarType(Handle)
    avarg(4) = CLngPtr(0) ' reserved
    avptr(4) = VarPtr(avarg(4))
    avt(5) = VarType(Handle)
    avarg(5) = VarPtr(m_pfnLoadAssembly) ' delegate
    avptr(5) = VarPtr(avarg(5))

    ' 4: CC_STDCALL
    e = DispCallFunc(0, pfnGetFunctionPointer, 4, vbLong, 6, avt(0), avptr(0), vr)
    If e < 0 Then
        Call hostfxr_close(Handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If
    e = vr
    If e < 0 Then
        Call LogErrorInfo(e)
        Call hostfxr_close(Handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If

    strMethodName = "LoadAssemblyFromFile"
    avarg(1) = StrPtr(strMethodName) ' method_name
    avarg(5) = VarPtr(m_pfnLoadAssemblyFromFile) ' delegate
    ' 4: CC_STDCALL
    e = DispCallFunc(0, pfnGetFunctionPointer, 4, vbLong, 6, avt(0), avptr(0), vr)
    If e < 0 Then
        Call hostfxr_close(Handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If
    e = vr
    If e < 0 Then
        Call LogErrorInfo(e)
        Call hostfxr_close(Handle)
        'Call FreeLibrary(hInstHostFXR)
        Call Err.Raise(e)
        Exit Sub
    End If

    m_handleHostFXR = Handle
    Call AddExitHandler(AddressOf OnExitHostFXR)
End Sub

Private Sub hostfxr_close(ByVal Handle As LongPtr)
    Dim avt() As Integer, avptr() As LongPtr, avarg() As Variant, vr As Variant
    Dim e As Long
    ReDim avt(0), avarg(0), avptr(0)
    avt(0) = VarType(Handle)
    avarg(0) = Handle
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
    ' Initializes vb2net
    Call InitializeVb2net

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
