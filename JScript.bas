Attribute VB_Name = "JScript"
'' ========================================================================= ''
' MIT License                                                                 '
'                                                                             '
' Copyright (c) 2020 Peter Donahue                                            '
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

'' ========================================================================= ''
'                                                                             '
'   Acknowledgements:                                                         '
'                                                                             '
'   These references were tremendously useful in understanding                '
'   COM and how to interact with it in Classic VB.                            '
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

Option Explicit

#Const ImplementRuntimeSourceResolveURL = True
#Const ImplementPublicJSONFunctions = True

#If Win32 Then
    Private Const LEN_PTR                       As Long = 4&
#Else
    Private Const LEN_PTR                       As Long = 8&
#End If

''
' Enums
'~~~~~~~~~~~~~~~~~~~
Private Enum ScriptItem
    IsVisible = &H2
    IsSource = &H4
    GlobalMembers = &H8
    IsPersistent = &H40
    CodeOnly = &H200
    NoCode = &H400
End Enum

Private Enum ScriptState
    Uninitialized = 0
    Started = 1
    Connected = 2
    Disconnected = 3
    Closed = 4
    Initialized = 5
End Enum

Private Enum ScriptText
    DelayExecution = &H1
    IsVisible = &H2
    IsExpression = &H20
    IsPersistent = &H40
    HostManagesSource = &H80
End Enum

Private Enum ScriptInfo
    IUnknown = 1
    ITypeInfo = 2
End Enum

''
' Custom Enums
'~~~~~~~~~~~~~~~~
Public Enum ScriptSourceType
#If False Then
    Dim Text, URL, Path
#End If
    Text
    URL
    Path
End Enum

Private Enum FileTypeEnum
    FileTypeAnsi = 1
    FileTypeUnicode
    FileTypeUtf8
    FileTypeUtf8NoBom
End Enum

''
' Private Types
'~~~~~~~~~~~~~~~~~~~~~~
Private Type ExceptionInfo
    wCode               As Integer
    wReserved           As Integer
    bstrSource          As String
    bstrDescription     As String
    bstrHelpFile        As String
    dwHelpContext       As Long
    pvReserved          As LongPtr
    pfnDeferredFillIn   As LongPtr
    hRes                As Long
End Type

Private Type Guid
    Data1               As Long
    Data2               As Integer
    Data3               As Integer
    Data4(7)            As Byte
End Type

Private Type MULTI_QI
    pIID                As LongPtr
    pItf                As LongPtr
    hr                  As Long
End Type

''
' GUIDS
'~~~~~~~~~~
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


Private Const S_OK                                  As Long = 0
Private Const CC_STDCALL                            As Long = 4&
Private Const CLSCTX_INPROC_SERVER                  As Long = 1&


Private Declare PtrSafe Function CoCreateInstanceEx Lib "ole32" (rclsid As Any, ByVal pUnkOuter As LongPtr, ByVal dwClsContext As Long, ByVal pServerInfo As LongPtr, ByVal dwCount As Long, rgmqResults As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As LongPtr, pCLSID As Any) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (ByVal rguid As LongPtr, ByVal lpsz As LongPtr, ByVal cchmax As Long) As Long
Private Declare PtrSafe Function IsBadCodePtr Lib "kernel32" (ByVal lpFn As LongPtr) As Long
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal offsetinVft As LongPtr, ByVal callconv As Long, ByVal retTYP As VbVarType, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As LongPtr, ByRef retVAR As Variant) As Long
Private Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare PtrSafe Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal sz As Long) As LongPtr
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pMem As LongPtr)
Private Declare PtrSafe Function lstrlenW Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropA" (ByVal HWnd As Long, ByVal lpString As String) As Long
Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropA" (ByVal HWnd As Long, ByVal lpString As String, ByVal hData As LongPtr) As Long

''
' FileRead & FileWrite
Private Declare PtrSafe Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long
Private Declare PtrSafe Function ApiCreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, ByVal lpSecurityAttributes As Long) As Long
''
' FileExists
Private Declare PtrSafe Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
'~~~~

Private Type ScriptExceptions
    Info            As ExceptionInfo
    SrcPosContext   As Long
    SrcPosLineNum   As Long
    SrcPosCharPos   As Long
    SourceLineText  As String
End Type

Private Type ScriptSiteVTables
    Site(0 To 10)      As LongPtr
    SiteDebug(0 To 10) As LongPtr
    SiteWindow(0 To 4) As LongPtr
End Type

Private Type ScriptHost
    Script              As MULTI_QI
    Parse               As MULTI_QI
    Site                As Long
    Debug               As LongPtr
    Window              As LongPtr
End Type


Private Host                    As ScriptHost
Private SiteVTables             As ScriptSiteVTables
Private LastException           As ScriptExceptions
Private ScriptObjects           As New Collection


Private Sub QuickTests()
    Dim hRes    As Long
    Dim RetVal  As Variant
    Static d       As Object
    Static D2      As Object
    
    
    ThisWorkbook.Save
    InitScriptHost
    
    Debug.Print Eval("1 + 202")
    'Set D2 = CreateObject("Scripting.Dictionary")
    Set d = CreateObject("Scripting.Dictionary")

    
    Debug.Print Eval("ThisWorkbook.Name")
    
    'Eval "var JDict = Dct;"
    
    Debug.Print Predicate(" d => d(""Mark"")")(d)
    Debug.Print Predicate(" d => d(""Mark"") = 25")(d)
    Debug.Print Predicate(" d => d(""Mark"")")(d)
    
    'Debug.Print d("Mark")
    Debug.Print Predicate(" d => d(""Peter"") = 1312")(d)
    Debug.Print Predicate(" d => d(""Peter"")")(d)
    
    Debug.Print Predicate(" d => d(""Peter"")")(d)
    Debug.Print Predicate(" d => d(""Mark"")")(d)
    
    
    Debug.Print Eval("var t = {'p':102, '4':'Mark'};")
    
    Debug.Print "->"; Fn("JSONStringify")(Eval("t"))
    
    Dim o As Object
    Set o = JsonParse("[1, 2,3,4,5]")
    
    Debug.Print Fn("JSONStringify")(o)
    Debug.Print JsonStringify(o)

    Eval "function testAddOne(x){ return x + 1;}"
    Debug.Print CallByName(IActiveScript_GetScriptDispatch(""), "testAddOne", VbMethod, 1)
End Sub



''
'
'   ## Public Functions
'
''

Public Function Eval(ByRef JScriptCode As String) As Variant
    If VBA.Trim$(JScriptCode) <> vbNullString Then
        Assign Eval, IActiveScript_ParseScriptText(JScriptCode)
    End If
End Function

Public Function Import(ByRef CodePathURL As String, Optional ByVal SourceType As ScriptSourceType = ScriptSourceType.Text)
    Assign Import, Eval(ResolveSourceType(CodePathURL, SourceType))
End Function

Public Function Fn(ByVal NameOrPredicate As String) As Object
    If InStr(NameOrPredicate, "=>") > 0 Then
        Set Fn = Predicate(NameOrPredicate)
    Else
        Set Fn = Eval(NameOrPredicate)
    End If
End Function

#If ImplementPublicJSONFunctions Then

    Public Property Get JsonParse(ByRef JsonText As String) As Object
        Set JsonParse = Eval("(function (){ return " & JsonText & "}())")
    End Property
    
    Public Property Get JsonStringify(ByRef JsonObject As Object) As String
        Static JsonStringifyFn As Object
        If JsonStringifyFn Is Nothing Then
            Set JsonStringifyFn = Fn("JSONStringify")
        End If
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
        If left(Parts(0), 1) <> "(" And right(Parts(0), 1) <> ")" Then
            Parts(0) = "(" & Parts(0) & ")"
        End If
    End If
    
    FullFunctionForm = "(function (){ return (function " & Parts(0) & " { return " & Parts(1) & "});} ())"
    Set Predicate = Eval(FullFunctionForm)
End Function


Private Function ResolveSourceType(ByRef CodePathURL As String, ByVal SourceType As ScriptSourceType)
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

Const VT_IACTIVESCRIPT_SETSCRIPTSITE As Long = &H3
Const VT_IACTIVESCRIPTPARSE_INITNEW  As Long = &H3
Dim ActiveScriptParseIID             As Guid
Dim ActiveScriptIID                  As Guid
Dim JScriptIID                       As Guid
Dim hRes                             As Long
    
    Call InitScriptSite
    If Host.Script.pItf = 0& Then
        CLSIDFromString ByVal StrPtr(sIID_IActiveScript), ActiveScriptIID
        CLSIDFromString ByVal StrPtr(sIID_IActiveScriptParse), ActiveScriptParseIID
        CLSIDFromString ByVal StrPtr(sIID_JScript9), JScriptIID
        With Host
            .Script.pIID = VarPtr(ActiveScriptIID)
            .Parse.pIID = VarPtr(ActiveScriptParseIID)
        End With
        Debug.Print CoCreateInstanceEx(JScriptIID, 0&, CLSCTX_INPROC_SERVER, 0&, 2&, ByVal VarPtr(Host))
        
        Debug.Assert CoCreateInstanceEx(JScriptIID, 0&, CLSCTX_INPROC_SERVER, 0&, 2&, ByVal VarPtr(Host)) = 0&
        Debug.Assert Invoke(Host.Script.pItf, VT_IACTIVESCRIPT_SETSCRIPTSITE * LEN_PTR, VarPtr(Host.Site)) = 0&
        Debug.Assert Invoke(Host.Parse.pItf, VT_IACTIVESCRIPTPARSE_INITNEW * LEN_PTR) = 0&
        Call InitScriptScope
        Call InitDefaultJsPackages
    End If
End Sub

Private Sub InitScriptSite()
    Const ACTIVE_SCRIPT_SITE_VTABLE_KEY As String = "ActiveScriptVTablePtr"
    
    'Is there an instance of the VTable currently in scope?
    If Host.Site = 0 Then
        
        'Was there an instance in scope at some point that we can restore?
        Host.Site = GetProp(Application.HWnd, ACTIVE_SCRIPT_SITE_VTABLE_KEY)
        
        'Allocate the memory
        If (Host.Site = 0) Then
            Host.Site = CoTaskMemAlloc(LenB(SiteVTables))
        End If
        
        'Grab/regrab the addresses of the methods -- since VBA recompiles
        'so frequently these can change out from under you
        With SiteVTables
            Dim pIUnknown_QueryInterface As LongPtr
            Dim pIUnknown_AddRef         As LongPtr
            Dim pIUnknown_Release        As LongPtr

            pIUnknown_QueryInterface = FnPtr(AddressOf IUnknown_QueryInterface)
            pIUnknown_AddRef = FnPtr(AddressOf IUnknown_AddRef)
            pIUnknown_Release = FnPtr(AddressOf IUnknown_Release)
            
            .Site(0) = pIUnknown_QueryInterface
            .Site(1) = pIUnknown_AddRef
            .Site(2) = pIUnknown_Release
            
            .Site(3) = FnPtr(AddressOf IActiveScriptSite_GetLCID)
            .Site(4) = FnPtr(AddressOf IActiveScriptSite_GetItemInfo)
            .Site(5) = FnPtr(AddressOf IActiveScriptSite_GetDocVersionString)
            .Site(6) = FnPtr(AddressOf IActiveScriptSite_OnScriptTerminate)
            .Site(7) = FnPtr(AddressOf IActiveScriptSite_OnStateChange)
            .Site(8) = FnPtr(AddressOf IActiveScriptSite_OnScriptError)
            .Site(9) = FnPtr(AddressOf IActiveScriptSite_OnEnterScript)
            .Site(10) = FnPtr(AddressOf IActiveScriptSite_OnLeaveScript)
            
            .SiteDebug(0) = pIUnknown_QueryInterface
            .SiteDebug(1) = pIUnknown_AddRef
            .SiteDebug(2) = pIUnknown_Release
            
            .SiteWindow(0) = pIUnknown_Release
            .SiteWindow(1) = pIUnknown_AddRef
            .SiteWindow(2) = pIUnknown_Release
            .SiteWindow(3) = FnPtr(AddressOf IActiveScriptSiteWindow_GetWindow)
            .SiteWindow(4) = FnPtr(AddressOf IActiveScriptSiteWindow_EnableModeless)
            
        End With
        
        'Store the most current method function pointers in the VTable
        CopyMemory ByVal Host.Site, ByVal VarPtr(SiteVTables), LenB(SiteVTables)
            
        'Store the VTable Address as a global application property this
        'ensure we arent reallocating one everything VBA recompiles,
        'this matters because i'm actually just letting this memory leak and
        'b/c i'm not calling CoTaskMemFree anywere.
        '
        ' That's fine if its only a few bytes -- and it will get cleared when the
        ' current excel application window is eventually closed
        
        Debug.Assert SetProp(Application.HWnd, ACTIVE_SCRIPT_SITE_VTABLE_KEY, Host.Site) <> 0
    End If
    Debug.Assert Host.Site <> 0
End Sub

Private Sub InitScriptScope()

    Dim Names     As Variant
    Dim Instances As Variant
    Dim Index     As Long
    
    Names = Array("Application", "ThisWorkbook")
    Instances = Array(Application, ThisWorkbook)
    
    For Index = LBound(Names) To UBound(Names)
        If Not CollectionContainsKey(ScriptObjects, Names(Index)) Then
            ScriptObjects.Add Instances(Index), Names(Index)
            IActiveScript_AddNamedItem Names(Index), ScriptItem.IsVisible Or ScriptItem.NoCode
        End If
    Next Index
    
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
    Eval "function JSONStringify(JsonObject) { return JSON.stringify(JsonObject); }"
    Eval "function JSONParse(JsonText) { return JSON.parse(JsonText); }"
    Eval "function IsArray(jsonObj) { return Object.prototype.toString.call(jsonObj) === '[object Array]';}"
    
End Sub


''
' IActiveScript Low-Level Wrappers
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Function IActiveScript_GetScriptDispatch(ByRef pStrItemName As String) As Object
Const VT_IACTIVESCRIPT_GETSCRIPTDISPATCH  As Long = &HA
Const FUNC_NAME                           As String = "IActiveScript_GetScriptDispatch"
Dim ParamCopies                           As Variant
Dim ToReturn                              As Object
Dim hRes                                  As Long
    InitScriptHost
    ParamCopies = Array(StrPtr(pStrItemName), VarPtr(ToReturn))
    hRes = Invoke(Host.Script.pItf, CLng(VT_IACTIVESCRIPT_GETSCRIPTDISPATCH * LEN_PTR), ParamCopies(0), ParamCopies(1))
    If hRes <> 0 Then Throw FUNC_NAME, ApiErrorText(hRes)
    Set IActiveScript_GetScriptDispatch = ToReturn
End Function

Private Sub IActiveScript_AddNamedItem(ByVal ObjectName As String, Optional Flags As Long = ScriptItem.IsVisible)
Const VT_IACTIVESCRIPT_ADDNAMEDITEM  As Long = &H8
Const FUNC_NAME                      As String = "IActiveScript_AddNamedItem"
Const VT_BYREF                       As Long = &H4000&
Dim Params(0 To 1)                   As Variant
Dim ObjNamePtr                       As LongPtr
Dim ObjNameStrPtr                    As LongPtr
    InitScriptHost
    Params(0) = ObjectName
    Params(1) = Flags
    CopyMemory ObjNamePtr, ByVal VarPtr(Params(0)), 2&
    If (ObjNamePtr And VT_BYREF) = 0& Then
        ObjNamePtr = VarPtr(Params(0)) + 8&
    Else
        CopyMemory ObjNamePtr, ByVal VarPtr(Params(0)) + 8&, LEN_PTR
    End If
    CopyMemory ObjNameStrPtr, ByVal ObjNamePtr, LEN_PTR
    Debug.Assert PointerToStringW(ObjNameStrPtr) <> vbNullString
    ThrowIfNotZero Invoke(Host.Script.pItf, CLng(VT_IACTIVESCRIPT_ADDNAMEDITEM * LEN_PTR), ObjNameStrPtr, Flags)
End Sub

Private Function IActiveScript_ParseScriptText(ByRef pstrCode As String, Optional ByRef pStrItemName As String, Optional ByRef pUnkContext As Object, Optional ByRef pstrDelimiter As String, Optional ByVal dwSourceContextCookie As Long, Optional ByVal ulStartingLineNumber As LongPtr, Optional ByVal dwFlags As ScriptText = ScriptText.IsExpression) As Variant

Const VT_IACTIVESCRIPT_PARSESCRIPTTEXT  As Long = &H5
Const FUNC_NAME                         As String = "IActiveScript_ParseScriptText"

Dim EvalResult                          As Variant
Dim hRes                                As Long

    InitScriptHost
    
    If Invoke(Host.Parse.pItf, CLng(VT_IACTIVESCRIPT_PARSESCRIPTTEXT * LEN_PTR), StrPtr(pstrCode), StrPtr(pStrItemName), ObjPtr(pUnkContext), StrPtr(pstrDelimiter), VarPtr(dwSourceContextCookie), VarPtr(ulStartingLineNumber), dwFlags, VarPtr(EvalResult), VarPtr(LastException)) <> 0& Then
        Call PrintErrorMessage
    End If
    
    Assign IActiveScript_ParseScriptText, EvalResult
    
End Function




''
' IActiveScriptSite Implementation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''
'   IUnknown Implementation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Function IUnknown_QueryInterface(ByVal pIUnknown As LongPtr, ByVal riid As LongPtr, ByRef pOut As LongPtr) As Long
    Const E_NOINTERFACE                         As Long = &H80004002
    Select Case GuidPtrString(riid)
    Case sIID_IUnknown, sIID_IActiveScriptSite
        CopyMemory pOut, ByVal VarPtr(Host.Site), LEN_PTR
    Case sIID_VBA_Debug_Mode
        IUnknown_QueryInterface = E_NOINTERFACE
        pOut = 0
    Case Else
        IUnknown_QueryInterface = E_NOINTERFACE
        pOut = 0
    End Select
End Function

Private Function IUnknown_AddRef(ByVal pIUnknown As LongPtr) As Long
'   Not Ref Counted
End Function

Private Function IUnknown_Release(ByVal pIUnknown As LongPtr) As Long
'   Not Ref Counted
End Function

''
'   IActiveScriptSite Method Handlers
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Function IActiveScriptSite_GetLCID(ByVal pIUnknown As LongPtr, ByVal plcid As LongPtr) As Long
    Const E_NOTIMPL  As Long = &H80004001
    CopyMemory ByVal plcid, E_NOTIMPL, 4
End Function

Private Static Function IActiveScriptSite_GetItemInfo(ByVal pIUnknown As LongPtr, ByVal pstrName As LongPtr, ByVal dwReturnMask As LongPtr, ByRef ppiunkItem As LongPtr, ByRef ppTi As LongPtr) As Long
    Const TYPE_E_ELEMENTNOTFOUND    As Long = &H8002802B
    Const FUNC_NAME                 As String = "IActiveScriptSite_GetItemInfo"
    Dim Name                        As String
    On Error GoTo CatchError
    Name = PointerToStringW(pstrName)
    If (dwReturnMask = ScriptInfo.IUnknown) Or (dwReturnMask = (ScriptInfo.ITypeInfo Or ScriptInfo.IUnknown)) Then
        ppiunkItem = ObjPtr(ScriptObjects(Name))
    End If
    If (dwReturnMask = ScriptInfo.ITypeInfo) Or (dwReturnMask = (ScriptInfo.ITypeInfo Or ScriptInfo.IUnknown)) Then
        ppTi = ObjPtr(GetTypeInformation(ScriptObjects(Name)))
    End If
    IActiveScriptSite_GetItemInfo = S_OK
Exit Function
CatchError:
    ppTi = 0&: ppiunkItem = 0&
    IActiveScriptSite_GetItemInfo = TYPE_E_ELEMENTNOTFOUND
End Function

Private Function IActiveScriptSite_GetDocVersionString(ByVal pIUnknown As LongPtr, ByVal pbstrVersionString As LongPtr) As Long
    IActiveScriptSite_GetDocVersionString = 0
End Function

Private Function IActiveScriptSite_OnScriptTerminate(ByVal pIUnknown As LongPtr, ByVal pvarResult As LongPtr, ByVal pExcepInfo As LongPtr) As Long
    IActiveScriptSite_OnScriptTerminate = 0
End Function

Private Function IActiveScriptSite_OnStateChange(ByVal pIUnknown As LongPtr, ByVal ssScriptState As LongPtr) As Long
    IActiveScriptSite_OnStateChange = 0&
End Function

Private Function IActiveScriptSite_OnScriptError(ByVal pIUnknown As LongPtr, ByVal pscripterror As LongPtr) As Long
Const VT_IACTIVESCRIPTERROR_GETEXCEPTIONINFO  As Long = &H3
Const VT_IACTIVESCRIPTERROR_GETSOURCEPOSITION As Long = &H4
Const VT_IACTIVESCRIPTERROR_GETSOURCELINETEXT As Long = &H5
    ThrowIfNotZero Invoke(ByVal pscripterror, ByVal (VT_IACTIVESCRIPTERROR_GETEXCEPTIONINFO * LEN_PTR), VarPtr(LastException.Info))
    ThrowIfNotZero Invoke(ByVal pscripterror, ByVal (VT_IACTIVESCRIPTERROR_GETSOURCEPOSITION * LEN_PTR), VarPtr(LastException.SrcPosContext), VarPtr(LastException.SrcPosLineNum), VarPtr(LastException.SrcPosCharPos))
    ThrowIfNotZero Invoke(ByVal pscripterror, ByVal (VT_IACTIVESCRIPTERROR_GETSOURCELINETEXT * LEN_PTR), VarPtr(LastException.SourceLineText))
End Function

Private Function IActiveScriptSite_OnEnterScript(ByVal pIUnknown As LongPtr) As Long
    Const VT_IUNKNOWN_ADDREF As Long = &H1
    DispCallFunc pIUnknown, CLng(VT_IUNKNOWN_ADDREF * LEN_PTR), 0&, 0&, 0&, 0&, 0&, 0&
    IActiveScriptSite_OnEnterScript = 0
End Function
Private Function IActiveScriptSite_OnLeaveScript(ByVal pIUnknown As LongPtr) As Long
    Const VT_IUNKNOWN_REMOVEREFF As Long = &H2
    DispCallFunc pIUnknown, CLng(VT_IUNKNOWN_REMOVEREFF * LEN_PTR), 0&, 0&, 0&, 0&, 0&, 0&
    IActiveScriptSite_OnLeaveScript = 0
End Function
Private Function IActiveScriptSiteWindow_GetWindow(ByVal pIUnknown As LongPtr, ByVal phwnd As LongPtr) As Long
    CopyMemory ByVal phwnd, ByVal VarPtr(Application.HWnd), LEN_PTR
    IActiveScriptSiteWindow_GetWindow = 0&
End Function
Private Function IActiveScriptSiteWindow_EnableModeless(ByVal pIUnknown As LongPtr, ByVal fEnable As LongPtr) As Long
    IActiveScriptSiteWindow_EnableModeless = 0&
End Function




''''
' Default JavaScript Code
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

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

    Const fromVBArrayCode = _
        "var fromVBArray = function (arrayIn) { " & _
        "   return new VBArray(arrayIn).toArray();" & _
        "}; "
        
    Const toVBArrayCode = _
        "var toVBArray = function (jsArray) { " & _
        "   var dict = new ActiveXObject('Scripting.Dictionary');" & _
        "   for (var i=0;i < jsArray.length; i++ ) dict.add(i,jsArray[i]);" & _
        "   return dict.items();" & _
        "};"
        
    VBArrayConversionImplementationCode = fromVBArrayCode & toVBArrayCode
    
End Property

Private Property Get RequireImplementationCode() As String

End Property



''
' Utility
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Function Invoke(ByVal ObjectPtr As LongPtr, ByVal FnOffset As LongPtr, ParamArray FunctionParameters() As Variant) As Variant
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                                                                                               '
    ' Derivative Of Function By VBForums user "LaVolpe"                                                             '
    ' See:  https://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)     '
    '                                                                                                               '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Const FUNC_NAME     As String = "Invoke"
Const CC_STDCALL    As Long = 4&

Dim pIndex          As Long
Dim pCount          As Long
Dim vParamPtr()     As LongPtr
Dim vParamType()    As Integer
Dim vRtn            As Variant
Dim vParams()       As Variant
    If UBound(FunctionParameters()) <> -1 Then
        vParams() = FunctionParameters()
        pCount = Abs(UBound(vParams) - LBound(vParams) + 1&)
    End If
    If pCount = 0& Then
        ReDim vParamPtr(0 To 0)
        ReDim vParamType(0 To 0)
    Else
        ReDim vParamPtr(0 To pCount - 1&)
        ReDim vParamType(0 To pCount - 1&)
        For pIndex = 0& To pCount - 1&
            vParamPtr(pIndex) = VarPtr(vParams(pIndex))
            vParamType(pIndex) = VarType(vParams(pIndex))
        Next
    End If
    pIndex = DispCallFunc(ObjectPtr, FnOffset, CC_STDCALL, vbLong, pCount, vParamType(0), vParamPtr(0), vRtn)
    If pIndex = 0& Then
        Assign Invoke, vRtn
    Else
        SetLastError pIndex
    End If
End Function

Private Function GetTypeInformation(ByRef AnObject As Object) As Object
Const CC_STDCALL        As Long = 4
Const GetTypeInfoVTO    As LongPtr = (4 * LEN_PTR)
Const SysDefaultLocale  As Long = &H800&
    Invoke ObjPtr(AnObject), GetTypeInfoVTO, 0&, SysDefaultLocale, VarPtr(GetTypeInformation)
End Function


Private Function CIUnknown(ByRef Obj As Object) As stdole.IUnknown
    Set CIUnknown = Obj
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

Private Function UnsignedAdd(ByVal Ptr As LongPtr, ByVal Offset As Long) As LongPtr

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                                                                     '
    ' I can't remember exactly whose iteration I borrowed here, but this technique        '
    ' is used by Matt Curland in "Advanced Visual Basic 6" and by Vladimir Vissoultchev   '
    ' (Github: https://github.com/wqweto) in numerous places.                             '
    '                                                                                     '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    
    UnsignedAdd = (((Ptr Xor &H80000000) + Offset) Xor &H80000000)
    
End Function

Private Function PointerToStringW(ByVal UnicodePointer As LongPtr) As String
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                                                                                               '
    ' Author: LaVolpe (VBForums Username)                                                                           '
    ' See:  https://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)     '
    '                                                                                                               '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    Dim StrLength   As Long
    If Not UnicodePointer = 0& Then
        StrLength = lstrlenW(UnicodePointer)
        If StrLength > 0& Then
            PointerToStringW = Space$(StrLength)
            CopyMemory ByVal StrPtr(PointerToStringW), ByVal UnicodePointer, StrLength * 2&
        End If
    End If
End Function

Private Function GuidPtrString(ByVal GuidPtr As LongPtr) As String
    GuidPtrString = String$(38, 0)
    StringFromGUID2 GuidPtr, StrPtr(GuidPtrString), 39&
End Function
Private Function GuidString(ByRef GuidValue() As Byte) As String
    GuidString = String$(38, 0)
    StringFromGUID2 VarPtr(GuidValue(0)), StrPtr(GuidString), 39&
End Function


Private Function FileRead(Filepath As String) As String

    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    '                                         '
    ' Author:  Vladimir Vissoultchev          '
    ' Github:  https://github.com/wqweto      '
    ' Project: VbPeg                          '
    ' Module:  VbPeg/src/mdMain.bas           '
    '                                         '
    ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
    
    Const ForReading    As Long = 1
    Const BOM_UTF       As String = "???"   '--- "\xEF\xBB\xBF"
    Const BOM_UNICODE   As String = "??"    '--- "\xFF\xFE"
    
    Dim lSize           As Long
    Dim sPrefix         As String
    Dim nFile           As Integer
    Dim sCharset        As String
    Dim oStream         As Object

    '--- get file size
    On Error GoTo EH
    If FileExists(Filepath) Then
        lSize = FileLen(Filepath)
    End If
    If lSize = 0 Then
        Exit Function
    End If
    '--- read first 50 chars
    nFile = FreeFile
    Open Filepath For Binary Access Read Shared As nFile
    sPrefix = String$(IIf(lSize < 50, lSize, 50), 0)
    Get nFile, , sPrefix
    Close nFile
    '--- figure out charset
    If left$(sPrefix, 3) = BOM_UTF Then
        sCharset = "UTF-8"
    ElseIf left$(sPrefix, 2) = BOM_UNICODE Or IsTextUnicode(ByVal sPrefix, Len(sPrefix), &HFFFF& - 2) <> 0 Then
        sCharset = "Unicode"
    ElseIf InStr(1, sPrefix, "<?xml", vbTextCompare) > 0 And InStr(1, sPrefix, "utf-8", vbTextCompare) > 0 Then
        '--- special xml encoding test
        sCharset = "UTF-8"
    End If
    '--- plain text: direct VB6 read
    If LenB(FileRead) = 0 And LenB(sCharset) = 0 Then
        nFile = FreeFile
        Open Filepath For Binary Access Read Shared As nFile
        FileRead = String$(lSize, 0)
        Get nFile, , FileRead
        Close nFile
    End If
    '--- plain text + unicode: use FileSystemObject
    If LenB(FileRead) = 0 And sCharset <> "UTF-8" Then
        On Error Resume Next  '--- checked
        FileRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(Filepath, ForReading, False, sCharset = "Unicode").ReadAll()
        On Error GoTo EH
    End If
    '--- plain text + unicode + utf-8: use ADODB.Stream
    If LenB(FileRead) = 0 Then
        Set oStream = CreateObject("ADODB.Stream")
        With oStream
            .Open
            If LenB(sCharset) <> 0 Then
                .Charset = sCharset
            End If
            .LoadFromFile Filepath
            FileRead = .ReadText()
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
    If GetFileAttributes(sFile) = INVALID_FILE_ATTRIBUTES Then
    Else
        FileExists = True
    End If
End Function

''
' Error Reporting
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private Sub ThrowIfNotZero(ByVal hRes As Long)
    If hRes <> 0& Then
        MsgBox ApiErrorText(hRes), vbOKOnly, "Error"
    End If
End Sub

Private Sub Throw(ByVal FuncName As String, ErrorMessage As String)
    MsgBox ErrorMessage, vbOKOnly, "Error: " & FuncName
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
    Dim Msg     As String
    Dim nRet    As Long
    
    Msg = Space$(1024)
    nRet = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrNum, 0&, Msg, Len(Msg), ByVal 0&)
    If nRet Then
        ApiErrorText = left$(Msg, nRet)
    Else
        ApiErrorText = "Error (" & ErrNum & ") not defined."
    End If
End Function

Private Function PrintErrorMessage() As String
    PrintErrorMessage = PrintErrorMessage & " -- Start: Exception Information -- " & vbNewLine
    With LastException
        With .Info
        PrintErrorMessage = PrintErrorMessage & "bstrDescription    : " & .bstrDescription & vbNewLine
        PrintErrorMessage = PrintErrorMessage & "bstrHelpFile       : " & .bstrHelpFile & vbNewLine
        PrintErrorMessage = PrintErrorMessage & "bstrSource         : " & .bstrSource & vbNewLine
        PrintErrorMessage = PrintErrorMessage & "dwHelpContext      : " & .dwHelpContext & vbNewLine
        PrintErrorMessage = PrintErrorMessage & "hRes               : " & .hRes & vbNewLine
        PrintErrorMessage = PrintErrorMessage & "pfnDeferredFillIn  : " & .pfnDeferredFillIn & vbNewLine
        PrintErrorMessage = PrintErrorMessage & "pvReserved         : " & .pvReserved & vbNewLine
        PrintErrorMessage = PrintErrorMessage & "wCode              : " & .wCode & vbNewLine
        PrintErrorMessage = PrintErrorMessage & "wReserved          : " & .wReserved & vbNewLine
        PrintErrorMessage = PrintErrorMessage & vbNewLine
        End With
        PrintErrorMessage = PrintErrorMessage & "SrcPosContext      : " & .SrcPosContext & vbNewLine
        PrintErrorMessage = PrintErrorMessage & "SrcPosLineNum      : " & .SrcPosLineNum & vbNewLine
        PrintErrorMessage = PrintErrorMessage & "SrcPosCharPos      : " & .SrcPosCharPos & vbNewLine
        PrintErrorMessage = PrintErrorMessage & vbNewLine
        PrintErrorMessage = PrintErrorMessage & "SourceLineText     :" & .SourceLineText & vbNewLine
    End With
    PrintErrorMessage = PrintErrorMessage & " -- End: Exception Infromation -- " & vbNewLine
    Debug.Print PrintErrorMessage
End Function


Private Function Assign(ByRef LHS As Variant, ByRef RHS As Variant) As Variant

    Const PROC_NAME         As String = "Assign"

    If IsObject(RHS) Then
        Set LHS = RHS
        Set Assign = LHS
    Else
        Let LHS = RHS
        Let Assign = LHS
    End If
End Function


Private Function NewLongPtrs(ParamArray LongPtrs() As Variant) As LongPtr()
    Dim ToReturn()  As LongPtr
    Dim Index       As Long
    If UBound(LongPtrs) > -1 Then
        ReDim Preserve ToReturn(0 To UBound(LongPtrs))
        For Index = LBound(LongPtrs) To UBound(LongPtrs)
            ToReturn(Index) = LongPtrs(Index)
        Next Index
    End If
    NewLongPtrs = ToReturn
End Function

Private Function NewIntegers(ParamArray Integers() As Variant) As Integer()
Dim ToReturn()  As Integer
Dim Index       As Long
    If UBound(Integers) > -1 Then
        ReDim Preserve ToReturn(0 To UBound(Integers))
        For Index = LBound(Integers) To UBound(Integers)
            ToReturn(Index) = Integers(Index)
        Next Index
    End If
    NewIntegers = ToReturn
End Function

Private Function CollectionContainsKey(ByRef Col As Object, ByVal Key As String)
    
    Const FUNC_NAME As String = "CollectionContainsKey"
    Dim Test        As Long

    On Error GoTo Catch
Try:
        Test = VarType(Col(Key)): CollectionContainsKey = True
        Exit Function
Catch:
        Select Case Err.Number
        Case 91:   CollectionContainsKey = False
        Case 5:    CollectionContainsKey = False
        Case Else: Throw FUNC_NAME, Err.Description
        End Select
        
End Function
