Attribute VB_Name = "JScript"

'' ========================================================================= ''
' JScript.bas — VBA host for the Windows JScript9 (ActiveScript) engine.
'' ========================================================================= ''

'' ========================================================================= ''
' MIT License                                                                 '
'                                                                             '
' Copyright (c) 2026 Peter Donahue                                            '
'                                                                             '
' Permission is hereby granted, free of charge, to any person obtaining a     '
' copy of this software and associated documentation files (the "Software"),  '
' to deal in the Software without restriction, including without limitation   '
' the rights to use, copy, modify, merge, publish, distribute, sublicense,    '
' and/or sell copies of the Software, and to permit persons to whom the       '
' Software is furnished to do so, subject to the following conditions:        '
'                                                                             '
' The above copyright notice and this permission notice shall be included in  '
' all copies or substantial portions of the Software.                         '
'                                                                             '
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR  '
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,    '
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL     '
' THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER  '
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING     '
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER         '
' DEALINGS IN THE SOFTWARE.                                                   '
'' ========================================================================= ''
'
'' ========================================================================= ''
'                                                                             '
'   General Acknowledgements:                                                 '
'                                                                             '
'   Author:  Matt Curland                                                     '
'   Book:    Advanced Visual Basic 6                                          '
'   See:     https://tinyurl.com/y2mghb93                                     '
'                                                                             '
'   Author:  Olaf Schmit (VBForums Username)                                  '
'   Project: vbFriendly Lightweight COM Interfaces                            '
'   See:     https://tinyurl.com/y5v4a2yr                                     '
'                                                                             '
'   Author:  LaVolpe (VBForums Username)                                      '
'   Project: FauxInterface                                                    '
'   See:     https://tinyurl.com/yxdxpe4o                                     '
'                                                                             '
'   Author:  David Zimmer (Github: dzzie)                                     '
'   Project: VB-ized IActiveScript Type Library                               '
'   See:     http://sandsprite.com                                            '
'                                                                             '
'' ========================================================================= ''

'' ========================================================================= ''
'  MAC GUARD                                                                  '
'  The Windows Scripting Host/JScript9 engine does not exist on macOS. Wr-    '
'  apping the entire module body prevents compile-time errors on the Mac VBA  '
'  runtime while still allowing the .bas file to be present in the project.   '
'' ========================================================================= ''
#If Not Mac Then

Option Explicit

' ---------------------------------------------------------------------------
' Compile-time feature flags
' ---------------------------------------------------------------------------
#Const ImplementRuntimeSourceResolveURL = True
#Const ImplementPublicJSONFunctions     = True


''
' Platform Compatibility 
'
' LongPtr polyfill for VBA6 (Office 2007 and earlier, always 32-bit).
'
' On VBA7 (Office 2010+), LongPtr is a native compiler type:
'   x86 build -> 4 bytes   x64 build -> 8 bytes
' On VBA6, LongPtr does not exist.  Declaring it as a Long-backed Enum
' (sizeof(Enum) = sizeof(Long) = 4) makes every LongPtr declaration below
' valid on VBA6 without any further conditional compile.
'
' Trick discovered by @Greedo — https://github.com/Greedquest
' See also: https://github.com/cristianbuse/VBA-MemoryTools/issues/3
#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If
' 
' LEN_PTR — byte-width of one pointer / vtable slot.
' Used throughout for vtable-offset arithmetic: slot_index * LEN_PTR.
'
' NOTE: Win32 is ALWAYS defined in VBA (even in 64-bit Office), so the
' original code's #If Win32 check was incorrect. Win64 is only defined
' in a genuine 64-bit Office build.
'
#If Win64 Then
    Private Const LEN_PTR As Long = 8&
#Else
    Private Const LEN_PTR As Long = 4&
#End If
'~~~~~~~~~~~~~~~~~~~


''
' Enums
'~~~~~~~~~~~~~~~~~~~
' ScriptItem flags passed to IActiveScript::AddNamedItem
Private Enum ScriptItem
    IsVisible       = &H2
    IsSource        = &H4
    GlobalMembers   = &H8
    IsPersistent    = &H40
    CodeOnly        = &H200
    NoCode          = &H400
End Enum

' SCRIPTSTATE values reported/requested via SetScriptState / GetScriptState
Private Enum ScriptState
    Uninitialized   = 0
    Started         = 1
    Connected       = 2
    Disconnected    = 3
    Closed          = 4
    Initialized     = 5
End Enum

' Flags for IActiveScriptParse::ParseScriptText
Private Enum ScriptText
    DelayExecution    = &H1
    IsVisible         = &H2
    IsExpression      = &H20
    IsPersistent      = &H40
    HostManagesSource = &H80
End Enum

' Masks for IActiveScriptSite::GetItemInfo dwReturnMask
Private Enum ScriptInfo
    IUnknown  = 1
    ITypeInfo = 2
End Enum

' SCRIPTTHREADSTATE returned by GetScriptThreadState
Private Enum ScriptThreadState
    NotInScript = 0
    Running     = 1
End Enum

' Public enum used by Import()
Public Enum ScriptSourceType
#If False Then
    Dim Text, URL, Path     ' suppress name-collision warnings
#End If
    Text
    URL
    Path
End Enum

Private Enum FileTypeEnum
    FileTypeAnsi    = 1
    FileTypeUnicode
    FileTypeUtf8
    FileTypeUtf8NoBom
End Enum

''
' Types
'
' ExceptionInfo — mirrors the COM EXCEPINFO struct.
'
' BSTRs are stored as raw LongPtr pointers rather than VBA String fields.
' This approach:
'   (a) avoids VBA's automatic BSTR reference counting which could
'       cause double-frees when the engine writes directly to memory;
'   (b) lets us add the explicit padding required on 64-bit without
'       relying on VBA's (absent) natural-alignment behaviour.
'
' VBA packs UDT fields consecutively with NO natural alignment padding.
' The C struct on x64 uses 8-byte natural alignment for BSTR/pointer fields,
' so we must insert explicit padding fields under Win64 to match the layout.
'
' Verified layouts:
'   x86 (32-bit): 2+2 + 4+4+4 + 4 + 4+4+4              = 32 bytes 
'   x64 (64-bit): 2+2+4 + 8+8+8 + 4+4 + 8+8+4+4        = 64 bytes 
'
Private Type ExceptionInfo
    wCode               As Integer      ' +0  (2 bytes)
    wReserved           As Integer      ' +2  (2 bytes)
#If Win64 Then
    Pad1                As Long         ' +4  (4 bytes — align BSTRs to 8-byte boundary)
#End If
    bstrSource          As LongPtr      ' +4  / +8   (BSTR pointer, ptr-sized)
    bstrDescription     As LongPtr      '            (BSTR pointer, ptr-sized)
    bstrHelpFile        As LongPtr      '            (BSTR pointer, ptr-sized)
    dwHelpContext       As Long         '            (DWORD, 4 bytes)
#If Win64 Then
    Pad2                As Long         '            (4 bytes — align pointer to 8-byte boundary)
#End If
    pvReserved          As LongPtr      '            (void*, ptr-sized)
    pfnDeferredFillIn   As LongPtr      '            (function pointer, ptr-sized)
    hRes                As Long         '            (SCODE/HRESULT, 4 bytes)
#If Win64 Then
    Pad3                As Long         '            (4 bytes — trailing alignment padding)
#End If
End Type

Private Type Guid
    Data1   As Long
    Data2   As Integer
    Data3   As Integer
    Data4(7) As Byte
End Type
'
' MULTI_QI — one query result for CoCreateInstanceEx.
' pIID and pItf are pointer-sized; hr is always a 32-bit HRESULT.
' The two MULTI_QI fields in ScriptHost must be adjacent so CoCreateInstanceEx
' can treat VarPtr(Host) as a MULTI_QI array[2].
Private Type MULTI_QI
    pIID    As LongPtr      ' [in]  Pointer to the requested IID
    pItf    As LongPtr      ' [out] Returned interface pointer
    hr      As Long         ' [out] Per-interface HRESULT (always 32-bit)
End Type
'
' ScriptHost — central state record.
'
' Script and Parse MUST be the first two fields and stay adjacent:
'   CoCreateInstanceEx receives VarPtr(Host) as a MULTI_QI array[2].
'
' Site / Debug / Window implement the faux-COM-object pattern:
'   *(VarPtr(Host.Site))   == Host.Site   == start of Site vtable   
'   *(VarPtr(Host.Window)) == Host.Window == start of Window vtable 
' so VarPtr(Host.Site) etc. are the COM object pointers.
'
Private Type ScriptHost
    Script  As MULTI_QI     ' IActiveScript interface (result from CoCreateInstanceEx)
    Parse   As MULTI_QI     ' IActiveScriptParse interface
    Site    As LongPtr      ' Points to the vtable block start  (= Site   vtable ptr)
    Debug   As LongPtr      ' Points into block at SiteDebug   sub-vtable
    Window  As LongPtr      ' Points into block at SiteWindow  sub-vtable
End Type

' Wraps raw EXCEPINFO with additional fields from IActiveScriptError
Private Type ScriptExceptions
    Info            As ExceptionInfo
    SrcPosContext   As Long
    SrcPosLineNum   As Long
    SrcPosCharPos   As Long
    SourceLineText  As String
End Type
'
' Contiguous vtable block written into CoTaskMem (Host.Site).
'
' Slot widths (bytes): LEN_PTR each.
' Byte offsets into the block:
'   Site(0..10)       11 slots  @ offset  0              -> Host.Site
'   SiteDebug(0..10)  11 slots  @ offset 11 * LEN_PTR    -> Host.Debug
'   SiteWindow(0..4)   5 slots  @ offset 22 * LEN_PTR    -> Host.Window
'
' For each sub-interface:
'   Host.Site/Debug/Window holds the block address for that sub-vtable.
'   VarPtr(Host.Site/Debug/Window) is the COM object pointer.
'   Deref once  -> Host.X  -> vtable array start
'   Deref twice -> slot[0] -> QueryInterface function pointer  
'
Private Type ScriptSiteVTables
    Site(0 To 10)       As LongPtr      ' IActiveScriptSite        (11 slots)
    SiteDebug(0 To 10)  As LongPtr      ' IActiveScriptSiteDebug   (11 slots)
    SiteWindow(0 To 4)  As LongPtr      ' IActiveScriptSiteWindow  ( 5 slots)
End Type
'~~~~~~~~~~~~~~~~~~~~

''
' GUIDS And Constants
'
Private Const sIID_IUnknown                         As String = "{00000000-0000-0000-C000-000000000046}"
Private Const sIID_IActiveScriptParse               As String = "{BB1A2AE2-A4F9-11CF-8F20-00805F2CD064}"
Private Const sIID_IActiveScriptError               As String = "{EAE1BA61-A4ED-11CF-8F20-00805F2CD064}"
Private Const sIID_IActiveScriptSiteWindow          As String = "{D10F6761-83E9-11CF-8F20-00805F2CD064}"
Private Const sIID_IActiveScriptSite                As String = "{DB01A1E3-A42B-11CF-8F20-00805F2CD064}"
Private Const sIID_IActiveScript                    As String = "{BB1A2AE1-A4F9-11CF-8F20-00805F2CD064}"
Private Const sIID_JScript9                         As String = "{16D51579-A30B-4C8B-A276-0FF4DC41E755}"
Private Const sIID_VBA_Debug_Mode                   As String = "{CACC1E85-622B-11D2-AA78-00C04F9901D2}"
Private Const sIID_IProvideClassInfo                As String = "{B196B283-BAB4-101A-B69C-00AA00341D07}"
Private Const sIID_IActiveScriptSiteDebug32         As String = "{51973C11-CB0C-11D0-B5C9-00A0244A0E7A}"
Private Const sIID_IActiveScriptSiteInterruptPoll   As String = "{539698A0-CDCA-11CF-A5EB-00AA0047A063}"
Private Const sIID_IOleCommandTarget                As String = "{B722BCCB-4E68-101B-A2BC-00AA00404770}"

Private Const S_OK                 As Long = 0&
Private Const E_NOTIMPL            As Long = &H80004001
Private Const E_NOINTERFACE        As Long = &H80004002
Private Const CC_STDCALL           As Long = 4&
Private Const CLSCTX_INPROC_SERVER As Long = 1&

' IActiveScript vtable indices (0=QI, 1=AddRef, 2=Release, then interface methods)
Private Const VTI_SETSCRIPTSITE         As Long = 3
Private Const VTI_GETSCRIPTSITE         As Long = 4
Private Const VTI_SETSCRIPTSTATE        As Long = 5
Private Const VTI_GETSCRIPTSTATE        As Long = 6
Private Const VTI_CLOSE                 As Long = 7
Private Const VTI_ADDNAMEDITEM          As Long = 8
Private Const VTI_ADDTYPELIB            As Long = 9
Private Const VTI_GETSCRIPTDISPATCH     As Long = 10
Private Const VTI_GETCURRENTTHREADID    As Long = 11
Private Const VTI_GETSCRIPTTHREADID     As Long = 12
Private Const VTI_GETSCRIPTTHREADSTATE  As Long = 13
Private Const VTI_INTERRUPTTHREAD       As Long = 14
Private Const VTI_CLONE                 As Long = 15

' IActiveScriptParse vtable indices
Private Const VTIP_INITNEW              As Long = 3
Private Const VTIP_ADDSCRIPTLET         As Long = 4
Private Const VTIP_PARSESCRIPTTEXT      As Long = 5

' IActiveScriptError vtable indices
Private Const VTIE_GETEXCEPTIONINFO     As Long = 3
Private Const VTIE_GETSOURCEPOSITION    As Long = 4
Private Const VTIE_GETSOURCELINETEXT    As Long = 5
'~~~~~~~~~~~~~~~~~~~~


''
' API Declarations
'
' Strategy: one #If VBA7 / #Else block for ALL declarations.
' Parameters are identical in both branches — every pointer-sized argument
' uses LongPtr, which the polyfill above resolves to Long on VBA6.
' The ONLY difference between the two branches is the PtrSafe keyword,
' which is required in 64-bit VBA7 and does not exist in VBA6.
'
#If VBA7 Then

    Private Declare PtrSafe Function CoCreateInstanceEx Lib "ole32" _
        (rclsid As Any, ByVal pUnkOuter As LongPtr, ByVal dwClsContext As Long, _
         ByVal pServerInfo As LongPtr, ByVal dwCount As Long, rgmqResults As LongPtr) As Long

    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

    Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
        (ByRef Destination As Any, ByVal Length As Long)

    Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" _
        (ByVal lpszProgID As LongPtr, pCLSID As Any) As Long

    Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" _
        (ByVal rguid As LongPtr, ByVal lpsz As LongPtr, ByVal cchmax As Long) As Long

    ' DispCallFunc: oVft is ULONG_PTR (pointer-sized); prgpvarg is LongPtr* (array of VarPtrs)
    Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" _
        (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal callconv As Long, _
         ByVal retTYP As VbVarType, ByVal paCNT As Long, ByRef paTypes As Integer, _
         ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long

    Private Declare PtrSafe Function SysAllocString Lib "oleaut32" _
        (ByVal pwsz As LongPtr) As LongPtr

    Private Declare PtrSafe Sub SysFreeString Lib "oleaut32" _
        (ByVal bstrString As LongPtr)

    Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
        (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
         ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
         Arguments As Long) As Long

    Private Declare PtrSafe Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)

    Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal sz As Long) As LongPtr

    Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pMem As LongPtr)

    Private Declare PtrSafe Function lstrlenW Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long

    Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropA" _
        (ByVal HWnd As LongPtr, ByVal lpString As String) As LongPtr

    Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropA" _
        (ByVal HWnd As LongPtr, ByVal lpString As String, ByVal hData As LongPtr) As Long

    Private Declare PtrSafe Function IsTextUnicode Lib "advapi32" _
        (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long

    Private Declare PtrSafe Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" _
        (ByVal lpFileName As String) As Long

#Else   ' VBA6 — no PtrSafe keyword; LongPtr resolves to Long via the Enum polyfill above

    Private Declare Function CoCreateInstanceEx Lib "ole32" _
        (rclsid As Any, ByVal pUnkOuter As LongPtr, ByVal dwClsContext As Long, _
         ByVal pServerInfo As LongPtr, ByVal dwCount As Long, rgmqResults As LongPtr) As Long

    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

    Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
        (ByRef Destination As Any, ByVal Length As Long)

    Private Declare Function CLSIDFromString Lib "ole32.dll" _
        (ByVal lpszProgID As LongPtr, pCLSID As Any) As Long

    Private Declare Function StringFromGUID2 Lib "ole32.dll" _
        (ByVal rguid As LongPtr, ByVal lpsz As LongPtr, ByVal cchmax As Long) As Long

    Private Declare Function DispCallFunc Lib "oleaut32" _
        (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal callconv As Long, _
         ByVal retTYP As VbVarType, ByVal paCNT As Long, ByRef paTypes As Integer, _
         ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long

    Private Declare Function SysAllocString Lib "oleaut32" _
        (ByVal pwsz As LongPtr) As LongPtr

    Private Declare Sub SysFreeString Lib "oleaut32" _
        (ByVal bstrString As LongPtr)

    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
        (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
         ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
         Arguments As Long) As Long

    Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)

    Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal sz As Long) As LongPtr

    Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pMem As LongPtr)

    Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long

    Private Declare Function GetProp Lib "user32" Alias "GetPropA" _
        (ByVal HWnd As LongPtr, ByVal lpString As String) As LongPtr

    Private Declare Function SetProp Lib "user32" Alias "SetPropA" _
        (ByVal HWnd As LongPtr, ByVal lpString As String, ByVal hData As LongPtr) As Long

    Private Declare Function IsTextUnicode Lib "advapi32" _
        (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long

    Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" _
        (ByVal lpFileName As String) As Long

#End If ' VBA7
'~~~~~~~~~~~~~~~~~~~~


''
' Module-Level State
'
Private Host            As ScriptHost
Private SiteVTables     As ScriptSiteVTables
Private LastException   As ScriptExceptions
Private ScriptObjects   As New Collection
'~~~~~~~~~~~~~~~~~~~~


''
'
'   ## Public Functions
'
''
' Evaluate a JScript expression or execute a statement.
' Returns the expression result as a Variant (Empty for statements).
Public Function Eval(ByRef JScriptCode As String) As Variant
    If VBA.Trim$(JScriptCode) <> vbNullString Then
        Assign Eval, IActiveScriptParse_ParseScriptText(JScriptCode)
    End If
End Function

' Load JScript from a code string, a file path, or a URL.
Public Function Import(ByRef CodePathURL As String, _
                       Optional ByVal SourceType As ScriptSourceType = ScriptSourceType.Text) As Variant
    Assign Import, Eval(ResolveSourceType(CodePathURL, SourceType))
End Function

' Return a named JScript function object, or convert an arrow expression
' (e.g. "x => x + 1") into a callable object.
Public Function Fn(ByVal NameOrPredicate As String) As Object
    If InStr(NameOrPredicate, "=>") > 0 Then
        Set Fn = Predicate(NameOrPredicate)
    Else
        Set Fn = Eval(NameOrPredicate)
    End If
End Function

' Expose a VBA object into the JScript global namespace.
Public Sub AddNamedObject(ByVal ObjectName As String, ByRef Obj As Object, _
                          Optional ByVal Flags As Long = 2)   ' 2 = IsVisible
    InitScriptHost
    If Not CollectionContainsKey(ScriptObjects, ObjectName) Then
        ScriptObjects.Add Obj, ObjectName
    End If
    IActiveScript_AddNamedItem ObjectName, Flags
End Sub

' Get/set the engine's current SCRIPTSTATE.
Public Property Get RunState() As ScriptState
    If Host.Script.pItf <> 0 Then RunState = IActiveScript_GetScriptState()
End Property
Public Property Let RunState(ByVal State As ScriptState)
    InitScriptHost
    IActiveScript_SetScriptState State
End Property

' Gracefully shut down the engine. Call from Workbook_BeforeClose.
Public Sub CloseScriptHost()
    On Error Resume Next
    If Host.Script.pItf <> 0 Then
        IActiveScript_SetScriptState ScriptState.Disconnected
        IActiveScript_Close
        ZeroMemory Host, LenB(Host)
    End If
    On Error GoTo 0
End Sub

#If ImplementPublicJSONFunctions Then

    Public Property Get JsonParse(ByRef JsonText As String) As Object
        Set JsonParse = Eval("(function(){ return " & JsonText & "}())")
    End Property

    Public Property Get JsonStringify(ByRef JsonObject As Object) As String
        Static JsonStringifyFn As Object
        If JsonStringifyFn Is Nothing Then Set JsonStringifyFn = Fn("JSONStringify")
        JsonStringify = JsonStringifyFn(JsonObject)
    End Property

#End If
''
'
'   ## End Public Functions
'
''


Private Function Predicate(ByRef ArrowExpression As String) As Object
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                                                         '
    ' Inspired by Series of Blog Posts by S. Meaden                           '
    ' See: http://exceldevelopmentplatform.blogspot.com/search/label/JScript  '
    '                                                                         '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    Dim Parts()          As String
    Dim FullFunctionForm As String
    Parts = Split(ArrowExpression, "=>")
    Debug.Assert UBound(Parts) = 1
    Parts(0) = VBA.Trim$(Parts(0))
    If Len(Parts(0)) > 0 Then
        If Left(Parts(0), 1) <> "(" And Right(Parts(0), 1) <> ")" Then
            Parts(0) = "(" & Parts(0) & ")"
        End If
    End If
    FullFunctionForm = "(function(){ return (function " & Parts(0) & " { return " & Parts(1) & "});} ())"
    Set Predicate = Eval(FullFunctionForm)
End Function

Private Function ResolveSourceType(ByRef CodePathURL As String, _
                                   ByVal SourceType As ScriptSourceType) As String
    Select Case SourceType
    Case ScriptSourceType.Text: ResolveSourceType = CodePathURL
    Case ScriptSourceType.Path: ResolveSourceType = FileRead(CodePathURL)
    Case ScriptSourceType.URL:
        #If ImplementRuntimeSourceResolveURL Then
            With CreateObject("MSXML2.XMLHTTP")
                .Open "GET", CodePathURL, False
                .Send
                ResolveSourceType = .responseText
            End With
        #End If
    End Select
End Function

''
' Initialization Procedures
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub InitScriptHost()
    Dim ActiveScriptParseIID    As Guid
    Dim ActiveScriptIID         As Guid
    Dim JScriptIID              As Guid
    Dim hRes                    As Long

    Call InitScriptSite

    If Host.Script.pItf = 0 Then
        CLSIDFromString ByVal StrPtr(sIID_IActiveScript),      ActiveScriptIID
        CLSIDFromString ByVal StrPtr(sIID_IActiveScriptParse), ActiveScriptParseIID
        CLSIDFromString ByVal StrPtr(sIID_JScript9),           JScriptIID

        Host.Script.pIID = VarPtr(ActiveScriptIID)
        Host.Parse.pIID  = VarPtr(ActiveScriptParseIID)

        ' Create both interfaces in one call. VarPtr(Host) is a MULTI_QI array[2]
        ' because Script and Parse are the first two adjacent fields in ScriptHost.
        hRes = CoCreateInstanceEx(JScriptIID, 0, CLSCTX_INPROC_SERVER, 0, 2&, ByVal VarPtr(Host))
        If hRes <> S_OK          Then Throw "InitScriptHost", "CoCreateInstanceEx failed: "    & ApiErrorText(hRes)
        If Host.Script.hr <> S_OK Then Throw "InitScriptHost", "IActiveScript QI failed: "     & ApiErrorText(Host.Script.hr)
        If Host.Parse.hr  <> S_OK Then Throw "InitScriptHost", "IActiveScriptParse QI failed: " & ApiErrorText(Host.Parse.hr)

        ' Hand our site to the engine.
        ' VarPtr(Host.Site) is the COM object ptr; deref -> Host.Site -> vtable.
        hRes = Invoke(Host.Script.pItf, VTI_SETSCRIPTSITE * LEN_PTR, VarPtr(Host.Site))
        If hRes <> S_OK Then Throw "InitScriptHost", "SetScriptSite failed: " & ApiErrorText(hRes)

        hRes = Invoke(Host.Parse.pItf, VTIP_INITNEW * LEN_PTR)
        If hRes <> S_OK Then Throw "InitScriptHost", "InitNew failed: " & ApiErrorText(hRes)

        ' Move to Connected so GetItemInfo callbacks fire for named items.
        IActiveScript_SetScriptState ScriptState.Connected

        Call InitScriptScope
        Call InitDefaultJsPackages
    End If
End Sub

Private Sub InitScriptSite()
    Const VTABLE_KEY As String = "ActiveScriptVTablePtr"
    
    'Is there an instance of the VTable currently in scope?
    If Host.Site = 0 Then
        'Was there an instance in scope at some point that we can restore?
        Host.Site = GetProp(Application.HWnd, VTABLE_KEY)

        If Host.Site = 0 Then
            Host.Site = CoTaskMemAlloc(LenB(SiteVTables))
            If Host.Site = 0 Then Throw "InitScriptSite", "CoTaskMemAlloc failed."
        End If

        
        ' Compute sub-vtable addresses at their fixed offsets inside the block.
        ' SiteDebug starts after Site(0..10) = 11 slots
        ' SiteWindow starts after SiteDebug  = 22 slots
        Host.Debug  = UnsignedAdd(Host.Site, 11 * LEN_PTR)
        Host.Window = UnsignedAdd(Host.Site, 22 * LEN_PTR)

        Dim pQI      As LongPtr: pQI      = FnPtr(AddressOf IUnknown_QueryInterface)
        Dim pAddRef  As LongPtr: pAddRef  = FnPtr(AddressOf IUnknown_AddRef)
        Dim pRelease As LongPtr: pRelease = FnPtr(AddressOf IUnknown_Release)
        
        ' Grab/regrab the addresses of the methods -- since VBA recompiles
        ' so frequently these can change out from under you
        With SiteVTables
            ' IActiveScriptSite vtable (indices 0-10)
            .Site(0)  = pQI
            .Site(1)  = pAddRef
            .Site(2)  = pRelease
            .Site(3)  = FnPtr(AddressOf IActiveScriptSite_GetLCID)
            .Site(4)  = FnPtr(AddressOf IActiveScriptSite_GetItemInfo)
            .Site(5)  = FnPtr(AddressOf IActiveScriptSite_GetDocVersionString)
            .Site(6)  = FnPtr(AddressOf IActiveScriptSite_OnScriptTerminate)
            .Site(7)  = FnPtr(AddressOf IActiveScriptSite_OnStateChange)
            .Site(8)  = FnPtr(AddressOf IActiveScriptSite_OnScriptError)
            .Site(9)  = FnPtr(AddressOf IActiveScriptSite_OnEnterScript)
            .Site(10) = FnPtr(AddressOf IActiveScriptSite_OnLeaveScript)

            ' IActiveScriptSiteDebug — QI returns E_NOINTERFACE; stubs only
            .SiteDebug(0) = pQI
            .SiteDebug(1) = pAddRef
            .SiteDebug(2) = pRelease
            ' Slots 3-10 remain zero (unreachable via our QI)

            ' IActiveScriptSiteWindow vtable (indices 0-4)
            .SiteWindow(0) = pQI        ' Slot 0 MUST be QueryInterface
            .SiteWindow(1) = pAddRef
            .SiteWindow(2) = pRelease
            .SiteWindow(3) = FnPtr(AddressOf IActiveScriptSiteWindow_GetWindow)
            .SiteWindow(4) = FnPtr(AddressOf IActiveScriptSiteWindow_EnableModeless)
        End With

        CopyMemory ByVal Host.Site, SiteVTables, LenB(SiteVTables)

        Debug.Assert SetProp(Application.HWnd, VTABLE_KEY, Host.Site) <> 0
    End If

    Debug.Assert Host.Site <> 0
End Sub

Private Sub InitScriptScope()
    Dim i As Long
    Dim Names()     As Variant: Names     = Array("Application", "ThisWorkbook")
    Dim Instances() As Variant: Instances = Array(Application, ThisWorkbook)
    For i = 0 To UBound(Names)
        If Not CollectionContainsKey(ScriptObjects, CStr(Names(i))) Then
            ScriptObjects.Add Instances(i), CStr(Names(i))
            IActiveScript_AddNamedItem CStr(Names(i)), ScriptItem.IsVisible Or ScriptItem.NoCode
        End If
    Next i
End Sub

Private Sub InitDefaultJsPackages()
    Eval JsonImplementationCode
    Eval VBArrayConversionImplementationCode
    Eval RequireImplementationCode
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                                                                                                   '
    ' JavaScript Helper Functions from Blogpost by S. Meadan                                                            '
    ' See: https://exceldevelopmentplatform.blogspot.com/2018/02/vba-jscript-passing-arrays-to-and-fro.html             '
    '                                                                                                                   '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    Eval "function JSONStringify(o)  { return JSON.stringify(o); }"
    Eval "function JSONParse(s)      { return JSON.parse(s); }"
    Eval "function IsArray(o)        { return Object.prototype.toString.call(o) === '[object Array]'; }"
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

''
' IActiveScript Low-Level Wrappers
'
'  VTable index:  
'   3 = SetScriptSite       4 = GetScriptSite          5 = SetScriptState  
'   6 = GetScriptState      7 = Close                  8 = AddNamedItem
'   9 = AddTypeLib         10 = GetScriptDispatch     11 = GetCurrentScriptThreadID
'  12 = GetScriptThreadID  13 = GetScriptThreadState  14 = InterruptScriptThread
'  15 = Clone
Private Function IActiveScript_GetScriptSite(ByRef riid As Guid) As LongPtr
    Dim pOut As LongPtr: Dim hRes As Long
    InitScriptHost
    hRes = Invoke(Host.Script.pItf, VTI_GETSCRIPTSITE * LEN_PTR, VarPtr(riid), VarPtr(pOut))
    If hRes <> S_OK Then Throw "IActiveScript_GetScriptSite", ApiErrorText(hRes)
    IActiveScript_GetScriptSite = pOut
End Function

Private Sub IActiveScript_SetScriptState(ByVal State As ScriptState)
    Dim hRes As Long
    hRes = Invoke(Host.Script.pItf, VTI_SETSCRIPTSTATE * LEN_PTR, CLng(State))
    If hRes <> S_OK Then Throw "IActiveScript_SetScriptState", ApiErrorText(hRes)
End Sub

Private Function IActiveScript_GetScriptState() As ScriptState
    Dim State As Long: Dim hRes As Long
    hRes = Invoke(Host.Script.pItf, VTI_GETSCRIPTSTATE * LEN_PTR, VarPtr(State))
    If hRes <> S_OK Then Throw "IActiveScript_GetScriptState", ApiErrorText(hRes)
    IActiveScript_GetScriptState = CInt(State)
End Function

Private Sub IActiveScript_Close()
    ThrowIfNotZero Invoke(Host.Script.pItf, VTI_CLOSE * LEN_PTR)
End Sub

' AddNamedItem: pass StrPtr directly — no need for fragile Variant traversal.
Private Sub IActiveScript_AddNamedItem(ByVal ObjectName As String, _
                                       Optional ByVal Flags As Long = 2)
    ThrowIfNotZero Invoke(Host.Script.pItf, VTI_ADDNAMEDITEM * LEN_PTR, _
                          StrPtr(ObjectName), Flags)
End Sub

Public Sub IActiveScript_AddTypeLib(ByRef rguidTypeLib As Guid, _
                                    ByVal dwMajor As Long, ByVal dwMinor As Long, _
                                    ByVal dwFlags As Long)
    InitScriptHost
    ThrowIfNotZero Invoke(Host.Script.pItf, VTI_ADDTYPELIB * LEN_PTR, _
                          VarPtr(rguidTypeLib), dwMajor, dwMinor, dwFlags)
End Sub

Public Function IActiveScript_GetScriptDispatch(ByRef pStrItemName As String) As Object
    Dim ToReturn As Object: Dim hRes As Long
    InitScriptHost
    hRes = Invoke(Host.Script.pItf, VTI_GETSCRIPTDISPATCH * LEN_PTR, _
                  StrPtr(pStrItemName), VarPtr(ToReturn))
    If hRes <> S_OK Then Throw "IActiveScript_GetScriptDispatch", ApiErrorText(hRes)
    Set IActiveScript_GetScriptDispatch = ToReturn
End Function

Public Function IActiveScript_GetCurrentScriptThreadID() As Long
    Dim ThreadID As Long: Dim hRes As Long
    InitScriptHost
    hRes = Invoke(Host.Script.pItf, VTI_GETCURRENTTHREADID * LEN_PTR, VarPtr(ThreadID))
    If hRes <> S_OK Then Throw "IActiveScript_GetCurrentScriptThreadID", ApiErrorText(hRes)
    IActiveScript_GetCurrentScriptThreadID = ThreadID
End Function

Public Function IActiveScript_GetScriptThreadID(ByVal dwWin32ThreadId As Long) As Long
    Dim stidThread As Long: Dim hRes As Long
    InitScriptHost
    hRes = Invoke(Host.Script.pItf, VTI_GETSCRIPTTHREADID * LEN_PTR, _
                  dwWin32ThreadId, VarPtr(stidThread))
    If hRes <> S_OK Then Throw "IActiveScript_GetScriptThreadID", ApiErrorText(hRes)
    IActiveScript_GetScriptThreadID = stidThread
End Function

Public Function IActiveScript_GetScriptThreadState(ByVal stidThread As Long) As ScriptThreadState
    Dim State As Long: Dim hRes As Long
    InitScriptHost
    hRes = Invoke(Host.Script.pItf, VTI_GETSCRIPTTHREADSTATE * LEN_PTR, _
                  stidThread, VarPtr(State))
    If hRes <> S_OK Then Throw "IActiveScript_GetScriptThreadState", ApiErrorText(hRes)
    IActiveScript_GetScriptThreadState = CInt(State)
End Function

Public Sub IActiveScript_InterruptScriptThread(ByVal stidThread As Long, _
                                               ByRef pExcepInfo As ExceptionInfo, _
                                               ByVal dwFlags As Long)
    InitScriptHost
    ThrowIfNotZero Invoke(Host.Script.pItf, VTI_INTERRUPTTHREAD * LEN_PTR, _
                          stidThread, VarPtr(pExcepInfo), dwFlags)
End Sub

' Returns a raw IActiveScript* pointer. Caller must Release when done.
Public Function IActiveScript_Clone() As LongPtr
    Dim pClone As LongPtr: Dim hRes As Long
    InitScriptHost
    hRes = Invoke(Host.Script.pItf, VTI_CLONE * LEN_PTR, VarPtr(pClone))
    If hRes <> S_OK Then Throw "IActiveScript_Clone", ApiErrorText(hRes)
    IActiveScript_Clone = pClone
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

''
' IActiveScriptParse Low-Level Wrappers
'
' VTable index:  
'  3 = InitNew   4 = AddScriptlet   5 = ParseScriptText
'
' dwSourceContextCookie and ulStartingLineNumber are [in] by-value DWORDs —
' passed directly, NOT via VarPtr (which was the original bug).
Private Function IActiveScriptParse_ParseScriptText( _
        ByRef pstrCode As String, _
        Optional ByRef pStrItemName As String = "", _
        Optional ByRef pUnkContext As Object = Nothing, _
        Optional ByRef pstrDelimiter As String = "", _
        Optional ByVal dwSourceContextCookie As Long = 0, _
        Optional ByVal ulStartingLineNumber As Long = 1, _
        Optional ByVal dwFlags As ScriptText = ScriptText.IsExpression) As Variant
    Dim EvalResult As Variant
    Dim hRes       As Long
    InitScriptHost
    hRes = Invoke(Host.Parse.pItf, VTIP_PARSESCRIPTTEXT * LEN_PTR, _
                  StrPtr(pstrCode), StrPtr(pStrItemName), ObjPtr(pUnkContext), _
                  StrPtr(pstrDelimiter), dwSourceContextCookie, ulStartingLineNumber, _
                  CLng(dwFlags), VarPtr(EvalResult), VarPtr(LastException))
    If hRes <> S_OK Then Call PrintErrorMessage
    Assign IActiveScriptParse_ParseScriptText, EvalResult
End Function

Public Function IActiveScriptParse_AddScriptlet( _
        ByRef pstrDefaultName As String, ByRef pstrCode As String, _
        ByRef pstrItemName As String, ByRef pstrSubItemName As String, _
        ByRef pstrEventName As String, ByRef pstrDelimiter As String, _
        Optional ByVal dwSourceContextCookie As Long = 0, _
        Optional ByVal ulStartingLineNumber As Long = 1, _
        Optional ByVal dwFlags As Long = 0) As String
    Dim pbstrName As LongPtr: Dim hRes As Long
    InitScriptHost
    hRes = Invoke(Host.Parse.pItf, VTIP_ADDSCRIPTLET * LEN_PTR, _
                  StrPtr(pstrDefaultName), StrPtr(pstrCode), _
                  StrPtr(pstrItemName), StrPtr(pstrSubItemName), _
                  StrPtr(pstrEventName), StrPtr(pstrDelimiter), _
                  dwSourceContextCookie, ulStartingLineNumber, dwFlags, _
                  VarPtr(pbstrName), VarPtr(LastException))
    If hRes <> S_OK Then Throw "IActiveScriptParse_AddScriptlet", ApiErrorText(hRes)
    If pbstrName <> 0 Then
        IActiveScriptParse_AddScriptlet = PointerToStringW(pbstrName)
        SysFreeString pbstrName
    End If
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

''
' IActiveScriptSite Low-Level Wrappers
'
' All callback parameters that carry COM pointers, BSTR pointers, or window
' handles are declared As LongPtr so they receive full-width values on x64.
'
''
'   IUnknown Implementation
'
Private Function IUnknown_QueryInterface(ByVal pSelf As LongPtr, ByVal riid As LongPtr, _
                                         ByRef pOut As LongPtr) As Long
    Select Case GuidPtrString(riid)
    Case sIID_IUnknown, sIID_IActiveScriptSite
        pOut = VarPtr(Host.Site)    ' object ptr: deref -> Host.Site -> vtable start
        IUnknown_QueryInterface = S_OK
    Case sIID_IActiveScriptSiteWindow
        pOut = VarPtr(Host.Window)  ' object ptr: deref -> Host.Window -> vtable start
        IUnknown_QueryInterface = S_OK
    Case Else
        pOut = 0
        IUnknown_QueryInterface = E_NOINTERFACE
    End Select
End Function

Private Function IUnknown_AddRef(ByVal pSelf As LongPtr) As Long
    IUnknown_AddRef = 1     ' not reference-counted
End Function

Private Function IUnknown_Release(ByVal pSelf As LongPtr) As Long
    IUnknown_Release = 1    ' not reference-counted
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

''
'   IActiveScriptSite Method Handlers
'
' Return E_NOTIMPL so the engine falls back to its default (system) locale.
Private Function IActiveScriptSite_GetLCID(ByVal pSelf As LongPtr, _
                                            ByVal plcid As LongPtr) As Long
    IActiveScriptSite_GetLCID = E_NOTIMPL
End Function

' dwReturnMask is a DWORD bit-field; test with And, not = (equality missed combined cases).
Private Static Function IActiveScriptSite_GetItemInfo(ByVal pSelf As LongPtr, _
                                                      ByVal pstrName As LongPtr, _
                                                      ByVal dwReturnMask As Long, _
                                                      ByRef ppiunkItem As LongPtr, _
                                                      ByRef ppTi As LongPtr) As Long
    Const TYPE_E_ELEMENTNOTFOUND As Long = &H8002802B
    Dim Name As String
    On Error GoTo CatchError
    Name = PointerToStringW(pstrName)
    If (dwReturnMask And ScriptInfo.IUnknown)  <> 0 Then ppiunkItem = ObjPtr(ScriptObjects(Name))
    If (dwReturnMask And ScriptInfo.ITypeInfo) <> 0 Then ppTi = ObjPtr(GetTypeInformation(ScriptObjects(Name)))
    IActiveScriptSite_GetItemInfo = S_OK
    Exit Function
CatchError:
    ppTi = 0: ppiunkItem = 0
    IActiveScriptSite_GetItemInfo = TYPE_E_ELEMENTNOTFOUND
End Function

' Allocate an empty BSTR "" and write its address to the output pointer.
Private Function IActiveScriptSite_GetDocVersionString(ByVal pSelf As LongPtr, _
                                                       ByVal pbstrVersionString As LongPtr) As Long
    Dim bstr As LongPtr
    bstr = SysAllocString(0)    ' NULL arg -> empty BSTR
    CopyMemory ByVal pbstrVersionString, bstr, LEN_PTR
    IActiveScriptSite_GetDocVersionString = S_OK
End Function

Private Function IActiveScriptSite_OnScriptTerminate(ByVal pSelf As LongPtr, _
                                                     ByVal pvarResult As LongPtr, _
                                                     ByVal pExcepInfo As LongPtr) As Long
    IActiveScriptSite_OnScriptTerminate = S_OK
End Function

Private Function IActiveScriptSite_OnStateChange(ByVal pSelf As LongPtr, _
                                                 ByVal ssScriptState As Long) As Long
    IActiveScriptSite_OnStateChange = S_OK
End Function

Private Function IActiveScriptSite_OnScriptError(ByVal pSelf As LongPtr, _
                                                 ByVal pScriptError As LongPtr) As Long
    ThrowIfNotZero Invoke(pScriptError, CLng(VTIE_GETEXCEPTIONINFO  * LEN_PTR), VarPtr(LastException.Info))
    ThrowIfNotZero Invoke(pScriptError, CLng(VTIE_GETSOURCEPOSITION * LEN_PTR), _
                          VarPtr(LastException.SrcPosContext), _
                          VarPtr(LastException.SrcPosLineNum), _
                          VarPtr(LastException.SrcPosCharPos))
    ThrowIfNotZero Invoke(pScriptError, CLng(VTIE_GETSOURCELINETEXT * LEN_PTR), _
                          VarPtr(LastException.SourceLineText))
    IActiveScriptSite_OnScriptError = S_OK
End Function

' Pure notification callbacks — returning S_OK is the entire correct implementation.
Private Function IActiveScriptSite_OnEnterScript(ByVal pSelf As LongPtr) As Long
    IActiveScriptSite_OnEnterScript = S_OK
End Function

Private Function IActiveScriptSite_OnLeaveScript(ByVal pSelf As LongPtr) As Long
    IActiveScriptSite_OnLeaveScript = S_OK
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

''
'   IActiveScriptSiteWindow Method Handlers
'
' Application.HWnd returns Long even on 64-bit VBA; widening to LongPtr
' zero-extends the value before CopyMemory writes LEN_PTR bytes.
Private Function IActiveScriptSiteWindow_GetWindow(ByVal pSelf As LongPtr, _
                                                   ByVal phwnd As LongPtr) As Long
    Dim hWndVal As LongPtr
    hWndVal = Application.HWnd     ' implicit widen: Long -> LongPtr (zero-extends on x64)
    CopyMemory ByVal phwnd, hWndVal, LEN_PTR
    IActiveScriptSiteWindow_GetWindow = S_OK
End Function

Private Function IActiveScriptSiteWindow_EnableModeless(ByVal pSelf As LongPtr, _
                                                        ByVal fEnable As Long) As Long
    IActiveScriptSiteWindow_EnableModeless = S_OK
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

''''
' Default JavaScript Code
'
Private Static Property Get JsonImplementationCode() As String
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                                               '
    ' Project: JSON2                                                '
    ' Author:  Douglas Crockford                                    '
    ' Source:  https://github.com/douglascrockford/JSON-js          '
    '                                                               '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    Const JSON2 As String = _
        "if(typeof JSON!==""object"")JSON={};(function(){var rx_one=/^[\],:{}\s]*$/;var rx_two=/\\(?:[""\\\/bfnrt]|u[0-9a-fA-F]{4})/g;var rx_three=/""[^""\\\n\r]*""|true|false|null|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?/g;var rx_four=/(?:^|:|,)(?:\s*\[)+/g;var rx_escapable=/[\\""\u0000-\u001f\u007f-\u009f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g;var rx_dangerous=/[\u0000\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]/g;function f(n){return n<10?""0""+" & _
        "n:n}function this_value(){return this.valueOf()}if(typeof Date.prototype.toJSON!==""function""){Date.prototype.toJSON=function(){return isFinite(this.valueOf())?this.getUTCFullYear()+""-""+f(this.getUTCMonth()+1)+""-""+f(this.getUTCDate())+""T""+f(this.getUTCHours())+"":""+f(this.getUTCMinutes())+"":""+f(this.getUTCSeconds())+""Z"":null};Boolean.prototype.toJSON=this_value;Number.prototype.toJSON=this_value;String.prototype.toJSON=this_value}var gap;var indent;var meta;var rep;function quote(string){rx_escapable.lastIndex=" & _
        "0;return rx_escapable.test(string)?'""'+string.replace(rx_escapable,function(a){var c=meta[a];return typeof c===""string""?c:""\\u""+(""0000""+a.charCodeAt(0).toString(16)).slice(-4)})+'""':'""'+string+'""'}function str(key,holder){var i;var k;var v;var length;var mind=gap;var partial;var value=holder[key];if(value&&typeof value===""object""&&typeof value.toJSON===""function"")value=value.toJSON(key);if(typeof rep===""function"")value=rep.call(holder,key,value);switch(typeof value){case ""string"":return quote(value);" & _
        "case ""number"":return isFinite(value)?String(value):""null"";case ""boolean"":case ""null"":return String(value);case ""object"":if(!value)return""null"";gap+=indent;partial=[];if(Object.prototype.toString.apply(value)===""[object Array]""){length=value.length;for(i=0;i<length;i+=1)partial[i]=str(i,value)||""null"";v=partial.length===0?""[]"":gap?""[\n""+gap+partial.join("",\n""+gap)+""\n""+mind+""]"":""[""+partial.join("","")+""]"";gap=mind;return v}if(rep&&typeof rep===""object""){length=rep.length;for(i=0;i<length;i+=1)if(typeof rep[i]===" & _
        """string""){k=rep[i];v=str(k,value);if(v)partial.push(quote(k)+(gap?"": "":"":"")+v)}}else for(k in value)if(Object.prototype.hasOwnProperty.call(value,k)){v=str(k,value);if(v)partial.push(quote(k)+(gap?"": "":"":"")+v)}v=partial.length===0?""{}"":gap?""{\n""+gap+partial.join("",\n""+gap)+""\n""+mind+""}"":""{""+partial.join("","")+""}"";gap=mind;return v}}if(typeof JSON.stringify!==""function""){meta={""\b"":""\\b"",""\t"":""\\t"",""\n"":""\\n"",""\f"":""\\f"",""\r"":""\\r"",'""':'\\""',""\\"":""\\\\""};JSON.stringify=function(value,replacer,space){var i;" & _
        "gap="""";indent="""";if(typeof space===""number"")for(i=0;i<space;i+=1)indent+="" "";else if(typeof space===""string"")indent=space;rep=replacer;if(replacer&&typeof replacer!==""function""&&(typeof replacer!==""object""||typeof replacer.length!==""number""))throw new Error(""JSON.stringify"");return str("""",{"""":value})}}if(typeof JSON.parse!==""function"")JSON.parse=function(text,reviver){var j;function walk(holder,key){var k;var v;var value=holder[key];if(value&&typeof value===""object"")for(k in value)if(Object.prototype.hasOwnProperty.call(value," & _
        "k)){v=walk(value,k);if(v!==undefined)value[k]=v;else delete value[k]}return reviver.call(holder,key,value)}text=String(text);rx_dangerous.lastIndex=0;if(rx_dangerous.test(text))text=text.replace(rx_dangerous,function(a){return""\\u""+(""0000""+a.charCodeAt(0).toString(16)).slice(-4)});if(rx_one.test(text.replace(rx_two,""@"").replace(rx_three,""]"").replace(rx_four,""""))){j=eval(""(""+text+"")"");return typeof reviver===""function""?walk({"""":j},""""):j}throw new SyntaxError(""JSON.parse"");}})();"
    JsonImplementationCode = JSON2
End Property

Private Property Get VBArrayConversionImplementationCode() As String
    Const fromVBArray As String = _
        "var fromVBArray=function(a){ return new VBArray(a).toArray(); };"
    Const toVBArray As String = _
        "var toVBArray=function(a){ var d=new ActiveXObject('Scripting.Dictionary');" & _
        "for(var i=0;i<a.length;i++) d.add(i,a[i]); return d.items(); };"
    VBArrayConversionImplementationCode = fromVBArray & toVBArray
End Property

Private Property Get RequireImplementationCode() As String
    Const requireJS As String = _
        "(function(g){var r={};function require(n){if(!r[n])throw new Error('Module not found: '+n);" & _
        "if(!r[n].e){var m={exports:{}};r[n].f(require,m,m.exports);r[n].e=m.exports;}return r[n].e;}" & _
        "require.define=function(n,f){r[n]={f:f}};g.require=require;}(this));"
    RequireImplementationCode = requireJS
End Property
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

''
' Utility
'
' Invoke — universal COM vtable dispatcher built on DispCallFunc.
'
'   ObjectPtr  = COM object pointer (engine reads *(ObjectPtr) to find vtable)
'   FnOffset   = byte offset of method slot from vtable start = index * LEN_PTR
'
' vParamPtr() stores VarPtr values (memory addresses), so it must be LongPtr:
'   32-bit: LongPtr = Long  (4 bytes, VarPtr returns Long)     
'   64-bit: LongPtr = LongLong (8 bytes, VarPtr returns LongLong) 
' DispCallFunc's paValues parameter is declared ByRef LongPtr, matching this.
'
Private Function Invoke(ByVal ObjectPtr As LongPtr, ByVal FnOffset As LongPtr, _
                        ParamArray FunctionParameters() As Variant) As Variant
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                                                                                               '
    ' Derivative Of Function By VBForums user "LaVolpe"                                                             '
    ' See:  https://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)     '
    '                                                                                                               '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

    Dim pCount       As Long
    Dim pIndex       As Long
    Dim vParamPtr()  As LongPtr     ' array of VarPtr values — must be LongPtr
    Dim vParamType() As Integer
    Dim vRtn         As Variant
    Dim vParams()    As Variant

    If UBound(FunctionParameters()) >= 0 Then
        vParams() = FunctionParameters()
        pCount = UBound(vParams) - LBound(vParams) + 1
    End If

    If pCount = 0 Then
        ReDim vParamPtr(0): ReDim vParamType(0)
    Else
        ReDim vParamPtr(0 To pCount - 1)
        ReDim vParamType(0 To pCount - 1)
        For pIndex = 0 To pCount - 1
            vParamPtr(pIndex)  = VarPtr(vParams(pIndex))   ' VarPtr -> LongPtr 
            vParamType(pIndex) = VarType(vParams(pIndex))
        Next pIndex
    End If

    pIndex = DispCallFunc(ObjectPtr, FnOffset, CC_STDCALL, vbLong, pCount, _
                          vParamType(0), vParamPtr(0), vRtn)
    If pIndex = S_OK Then
        Assign Invoke, vRtn
    Else
        SetLastError pIndex
    End If
End Function

Private Function GetTypeInformation(ByRef AnObject As Object) As Object
    ' IDispatch vtable index 4 = GetTypeInfo; LOCALE_SYSTEM_DEFAULT = &H800
    Const GetTypeInfoOffset As Long = 4 * LEN_PTR  ' 16 (x86) or 32 (x64)
    Const SysDefaultLocale  As Long = &H800&
    Invoke ObjPtr(AnObject), GetTypeInfoOffset, 0&, SysDefaultLocale, VarPtr(GetTypeInformation)
End Function

Private Function FnPtr(ByVal Address As LongPtr) As LongPtr
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                                                   '
    ' This is a classic VB Trick for getting the address of a function. '
    ' I didnt discover it independantly, but also wouldnt know where to '
    ' attribute it at this point                                        '
    '                                                                   '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    FnPtr = Address
End Function

' ---------------------------------------------------------------------------
' UnsignedAdd — pointer addition that avoids signed-integer overflow.
'
' The XOR trick flips the sign bit, does the addition in unsigned space,
' then restores it. The sign-bit mask must match the width of LongPtr:
'   32-bit: &H80000000      (Long literal)
'   64-bit: &H8000000000000000^  (LongLong literal, only valid under Win64)
' ---------------------------------------------------------------------------
Private Function UnsignedAdd(ByVal Ptr As LongPtr, ByVal Offset As Long) As LongPtr
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                                                                     '
    ' I can't remember exactly whose iteration I borrowed here, but this technique        '
    ' is used by Matt Curland in "Advanced Visual Basic 6" and by Vladimir Vissoultchev   '
    ' (Github: https://github.com/wqweto) in numerous places.                             '
    '                                                                                     '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
#If Win64 Then
    UnsignedAdd = (((Ptr Xor &H8000000000000000^) + CLngLng(Offset)) Xor &H8000000000000000^)
#Else
    UnsignedAdd = (((Ptr Xor &H80000000) + Offset) Xor &H80000000)
#End If
End Function

Private Function PointerToStringW(ByVal UnicodePointer As LongPtr) As String
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                                                                                               '
    ' Author: LaVolpe (VBForums Username)                                                                           '
    ' See:  https://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)     '
    '                                                                                                               '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    Dim StrLength As Long
    If UnicodePointer <> 0 Then
        StrLength = lstrlenW(UnicodePointer)
        If StrLength > 0 Then
            PointerToStringW = Space$(StrLength)
            CopyMemory ByVal StrPtr(PointerToStringW), ByVal UnicodePointer, StrLength * 2
        End If
    End If
End Function

Private Function GuidPtrString(ByVal GuidPtr As LongPtr) As String
    GuidPtrString = String$(38, 0)
    StringFromGUID2 GuidPtr, StrPtr(GuidPtrString), 39
End Function

Private Function GuidString(ByRef GuidValue() As Byte) As String
    GuidString = String$(38, 0)
    StringFromGUID2 VarPtr(GuidValue(0)), StrPtr(GuidString), 39
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

''
' Error Reporting
'
Private Sub ThrowIfNotZero(ByVal hRes As Long)
    If hRes <> S_OK Then MsgBox ApiErrorText(hRes), vbOKOnly Or vbExclamation, "Script Error"
End Sub

Private Sub Throw(ByVal FuncName As String, ByVal ErrorMessage As String)
    MsgBox ErrorMessage, vbOKOnly Or vbCritical, "Error in " & FuncName
    End
End Sub

Private Function ApiErrorText(ByVal ErrNum As Long) As String
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                         '
    ' Author:  Vladimir Vissoultchev          '
    ' Github:  https://github.com/wqweto      '
    ' Project: UcsFiscalPrinters              '
    ' Module:  src/Shared/mdGlobals.bas       '
    '                                         '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

    Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
    Dim Msg  As String: Msg = Space$(1024)
    Dim nRet As Long:   nRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrNum, 0&, Msg, Len(Msg), ByVal 0&)
    If nRet > 0 Then ApiErrorText = Left$(Msg, nRet) Else ApiErrorText = "HRESULT 0x" & Hex$(ErrNum)
End Function
'
' PrintErrorMessage — reads BSTR fields from ExceptionInfo via PointerToStringW
' because bstrSource etc. are stored as raw LongPtr (not VBA Strings).
' Frees the engine-allocated BSTRs after reading to avoid memory leaks.
'
Private Function PrintErrorMessage() As String
    Dim src  As String: src  = PointerToStringW(LastException.Info.bstrSource)
    Dim desc As String: desc = PointerToStringW(LastException.Info.bstrDescription)
    Dim help As String: help = PointerToStringW(LastException.Info.bstrHelpFile)

    ' Free engine-allocated BSTRs now that we have copied their content.
    If LastException.Info.bstrSource      <> 0 Then SysFreeString LastException.Info.bstrSource:      LastException.Info.bstrSource = 0
    If LastException.Info.bstrDescription <> 0 Then SysFreeString LastException.Info.bstrDescription: LastException.Info.bstrDescription = 0
    If LastException.Info.bstrHelpFile    <> 0 Then SysFreeString LastException.Info.bstrHelpFile:    LastException.Info.bstrHelpFile = 0

    PrintErrorMessage = " -- Script Exception --"              & vbNewLine & _
                        "Source      : " & src                 & vbNewLine & _
                        "Description : " & desc                & vbNewLine & _
                        "HelpFile    : " & help                & vbNewLine & _
                        "HelpContext : " & LastException.Info.dwHelpContext & vbNewLine & _
                        "HRESULT     : 0x" & Hex$(LastException.Info.hRes) & vbNewLine & _
                        "Line        : " & LastException.SrcPosLineNum     & vbNewLine & _
                        "Char        : " & LastException.SrcPosCharPos     & vbNewLine & _
                        "Source line : " & LastException.SourceLineText    & vbNewLine

    Debug.Print PrintErrorMessage
    MsgBox PrintErrorMessage, vbOKOnly Or vbExclamation, "JScript Error"
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

''
' General Utility
'
Private Function Assign(ByRef LHS As Variant, ByRef RHS As Variant) As Variant
    If IsObject(RHS) Then: Set LHS = RHS:  Set Assign = LHS
    Else:                      LHS = RHS:      Assign = LHS
    End If
End Function

Private Function CollectionContainsKey(ByRef Col As Object, ByVal Key As String) As Boolean
    Dim Test As Long
    On Error GoTo Catch
    Test = VarType(Col(Key)):  CollectionContainsKey = True:  Exit Function
Catch:
    Select Case Err.Number
    Case 5, 91: CollectionContainsKey = False
    Case Else:  Throw "CollectionContainsKey", Err.Description
    End Select
End Function
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


''
' File I/O
'
Private Function FileRead(Filepath As String) As String
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                         '
    ' Author:  Vladimir Vissoultchev          '
    ' Github:  https://github.com/wqweto      '
    ' Project: VbPeg                          '
    ' Module:  VbPeg/src/mdMain.bas           '
    '                                         '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    Const ForReading  As Long = 1
    Const BOM_UTF     As String = "???"
    Const BOM_UNICODE As String = "??"
    Dim lSize As Long: Dim sPrefix As String: Dim nFile As Integer
    Dim sCharset As String: Dim oStream As Object
    On Error GoTo EH
    If FileExists(Filepath) Then lSize = FileLen(Filepath)
    If lSize = 0 Then Exit Function
    nFile = FreeFile
    Open Filepath For Binary Access Read Shared As nFile
    sPrefix = String$(IIf(lSize < 50, lSize, 50), 0)
    Get nFile, , sPrefix: Close nFile
    If Left$(sPrefix, 3) = BOM_UTF Then
        sCharset = "UTF-8"
    ElseIf Left$(sPrefix, 2) = BOM_UNICODE Or IsTextUnicode(ByVal sPrefix, Len(sPrefix), &HFFFF& - 2) <> 0 Then
        sCharset = "Unicode"
    ElseIf InStr(1, sPrefix, "<?xml", vbTextCompare) > 0 And InStr(1, sPrefix, "utf-8", vbTextCompare) > 0 Then
        sCharset = "UTF-8"
    End If
    If LenB(FileRead) = 0 And LenB(sCharset) = 0 Then
        nFile = FreeFile: Open Filepath For Binary Access Read Shared As nFile
        FileRead = String$(lSize, 0): Get nFile, , FileRead: Close nFile
    End If
    If LenB(FileRead) = 0 And sCharset <> "UTF-8" Then
        On Error Resume Next
        FileRead = CreateObject("Scripting.FileSystemObject") _
                       .OpenTextFile(Filepath, ForReading, False, sCharset = "Unicode").ReadAll()
        On Error GoTo EH
    End If
    If LenB(FileRead) = 0 Then
        Set oStream = CreateObject("ADODB.Stream")
        With oStream
            .Open
            If LenB(sCharset) <> 0 Then .Charset = sCharset
            .LoadFromFile Filepath: FileRead = .ReadText()
        End With
    End If
    Exit Function
EH:
End Function

Private Function FileExists(sFile As String) As Boolean
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                         '
    ' Author:  Vladimir Vissoultchev          '
    ' Github:  https://github.com/wqweto      '
    ' Project: VbPeg                          '
    ' Module:  VbPeg/src/mdMain.bas           '
    '                                         '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    Const INVALID_FILE_ATTRIBUTES As Long = -1
    FileExists = (GetFileAttributes(sFile) <> INVALID_FILE_ATTRIBUTES)
End Function

#End If ' Not Mac
