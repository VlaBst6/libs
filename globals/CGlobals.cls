VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGlobals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Source As Long, ByVal length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Declare Sub DebugBreak Lib "kernel32" ()
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function CallAsm Lib "user32" Alias "CallWindowProcA" (ByRef lpBytes As Any, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallAsmAddr Lib "user32" Alias "CallWindowProcA" (ByVal lpCode As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal length As Long)


'all but jet are odbc jet is ole
Public Enum dbServers
    Access
    JetAccess2k
    MsSql2k
    mysql
    DSN
    FileDsn
    dBase
End Enum


Property Get LANG_US() As Long
    LANG_US = &H409
End Property

Public Function isIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 / 0
    isIde = False
    Exit Function
hell:
    isIde = True
End Function

Function TopMost(frm As Object, Optional ontop As Boolean = True)
    On Error Resume Next
    s = IIf(ontop, HWND_TOPMOST, HWND_NOTOPMOST)
    SetWindowPos frm.hWnd, s, frm.Left / 15, frm.Top / 15, frm.Width / 15, frm.Height / 15, 0
End Function

Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 67
    Dim lResult As Long
    Dim iCreated As Boolean
    
    'the path must actually exist to get the short path name !!
    If Not FileExists(sFile) Then
        iCreated = True
        writeFile sFile, ""
    End If
    
    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, _
    Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)
    
    If Len(GetShortName) = 0 Then GetShortName = sFile
    
    If iCreated Then Kill sFile
    
End Function

Public Function bHexDump(b() As Byte, Optional hexOnly = 0) As String
    Dim tmp As String
    Dim b2() As Byte
    tmp = StrConv(b, vbUnicode, LANG_US) 'this is lazy, but hexdump uses 1 based ary internally so whatever..
    bHexDump = HexDump(tmp, hexOnly)
End Function

Function HexDump(ByVal str, Optional hexOnly = 0) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
    Dim offset As Long
    
    offset = 0
    str = " " & str
    ary = StrConv(str, vbFromUnicode, LANG_US)
    
    chars = "   "
    For i = 1 To UBound(ary)
        tt = Hex(ary(i))
        If Len(tt) = 1 Then tt = "0" & tt
        tmp = tmp & tt & " "
        x = ary(i)
        'chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".") 'x > 191 causes \x0 problems on non us systems... asc(chr(x)) = 0
        chars = chars & IIf((x > 32 And x < 127), Chr(x), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            If hexOnly = 0 Then
                push s, h & "   " & tmp & chars
            Else
                push s, tmp
            End If
            offset = offset + 16
            tmp = Empty
            chars = "   "
        End If
    Next
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        If hexOnly = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            h = h & "   " & tmp
            While Len(h) <= 56: h = h & " ": Wend
            push s, h & chars
        Else
            push s, tmp
        End If
    End If
    
    HexDump = Join(s, vbCrLf)
    
    If hexOnly <> 0 Then
        HexDump = Replace(HexDump, " ", "")
        HexDump = Replace(HexDump, vbCrLf, "")
    End If
    
End Function

Sub FormPos(fform As Object, Optional andSize As Boolean = False, Optional save_mode As Boolean = False)
    
    On Error Resume Next
    
    Dim f, sz
    f = Split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For i = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(i), VbGet)
            SaveSetting App.EXEName, fform.Name & ".FormPos", f(i), ff
        Else
            def = CallByName(fform, f(i), VbGet)
            ff = GetSetting(App.EXEName, fform.Name & ".FormPos", f(i), def)
            CallByName fform, f(i), VbLet, ff
        End If
    Next
    
End Sub

Sub SaveMySetting(key, value)
    SaveSetting App.EXEName, "Settings", key, value
End Sub

Function GetMySetting(key, Optional defaultval = "")
    GetMySetting = GetSetting(App.EXEName, "Settings", key, defaultval)
End Function

Function FolderExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbDirectory) <> "" Then FolderExists = True _
  Else FolderExists = False
End Function

Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then
     If Err.Number <> 0 Then Exit Function
     FileExists = True
  End If
End Function

Function RandomNum() As Long
    Dim tmp As Long
    Dim tries As Long
    
    On Error Resume Next

    Do While 1
        Err.Clear
        Randomize
        tmp = Round(Timer * Now * Rnd(), 0)
        RandomNum = tmp
        If Err.Number = 0 Then Exit Function
        If tries < 100 Then
            tries = tries + 1
        Else
            Exit Do
        End If
    Loop
    
    RandomNum = GetTickCount
    
End Function

Function ReadFile(filename) As String 'this one should be binary safe...
  On Error GoTo hell
  f = FreeFile
  Dim b() As Byte
  Open filename For Binary As #f
  ReDim b(LOF(f) - 1)
  Get f, , b()
  Close #f
  ReadFile = StrConv(b(), vbUnicode, LANG_US)
  Exit Function
hell:   ReadFile = ""
End Function

Function writeFile(path, it) As Boolean 'this one should be binary safe...
    On Error GoTo hell
    Dim b() As Byte
    If FileExists(path) Then Kill path
    f = FreeFile
    b() = StrConv(it, vbFromUnicode, LANG_US)
    Open path For Binary As #f
    Put f, , b()
    Close f
    writeFile = True
    Exit Function
hell: writeFile = False
End Function

Sub AppendFile(path, it) 'not binary safe
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Function pop(ary)
    On Error GoTo isEmpty
    x = UBound(ary)
    pop = ary(x)
    ReDim Preserve ary(x - 1)
    Exit Function
isEmpty: Erase ary
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function keyExists(t As String, c As Collection) As Boolean
    On Error GoTo hell
    x = c(t)
    keyExists = True
    Exit Function
hell:
End Function

Function AnyofTheseInstr(data, match, Optional compare As VbCompareMethod = vbTextCompare) As Boolean
Attribute AnyofTheseInstr.VB_Description = "Accepts comma seperated value list of strings to look for"
    Dim tmp() As String
    Dim x
    tmp = Split(match, ",")
    For Each x In tmp
        If InStr(1, data, x, compare) > 0 Then
            AnyofTheseInstr = True
            Exit Function
        End If
    Next
End Function

Function GetCount(str, what) 'as long
    On Error Resume Next
    GetCount = UBound(Split(str, what)) + 1
    If Len(GetCount) = 0 Then GetCount = 0
End Function

Public Function isHexNum(v) As Boolean
    On Error Resume Next
    x = CLng("&h" & v)
    If Err.Number = 0 Then isHexNum = True
    Err.Clear
End Function

Public Function FileSize(fPath As String) As String
    Dim fsize As Long
    Dim szName As String
    On Error GoTo hell
    
    fsize = FileLen(fPath)
    
    szName = " bytes"
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Kb"
    End If
    
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Mb"
    End If
    
    FileSize = fsize & szName
    
    Exit Function
hell:
    
End Function

Public Function GetFileVersion(Optional ByVal PathWithFilename As String) As String
    ' return file-properties of given file  (EXE , DLL , OCX)
    'http://support.microsoft.com/default.aspx?scid=kb;en-us;160042
    
    If Len(PathWithFilename) = 0 Then Exit Function
    
    Dim lngBufferlen As Long
    Dim lngDummy As Long
    Dim lngRc As Long
    Dim lngVerPointer As Long
    Dim lngHexNumber As Long
    Dim b() As Byte
    Dim b2() As Byte
    Dim strBuffer As String
    Dim strLangCharset As String
    Dim strTemp As String
    Dim n As Long
    
    ReDim b2(500)
    
    lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
    If lngBufferlen <= 0 Then Exit Function
    
    ReDim b(lngBufferlen)
    lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, b(0))
    If lngRc = 0 Then Exit Function
    
    lngRc = VerQueryValue(b(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen)
    If lngRc = 0 Then Exit Function
    
    MoveMemory b2(0), lngVerPointer, lngBufferlen
    lngHexNumber = b2(2) + b2(3) * &H100 + b2(0) * &H10000 + b2(1) * &H1000000
    strLangCharset = Right("0000000" & Hex(lngHexNumber), 8)
    
    strBuffer = String$(800, 0)
    strTemp = "\StringFileInfo\" & strLangCharset & "\FileVersion"
    lngRc = VerQueryValue(b(0), strTemp, lngVerPointer, lngBufferlen)
    If lngRc = 0 Then Exit Function
    
    lstrcpy strBuffer, lngVerPointer
    n = InStr(strBuffer, Chr(0)) - 1
    If n > 0 Then
        strBuffer = Mid$(strBuffer, 1, n)
        GetFileVersion = strBuffer
    End If
   
End Function

Public Sub LV_ColumnSort(ListViewControl As Object, Column As Object)
    On Error Resume Next
    With ListViewControl
       If .SortKey <> Column.Index - 1 Then
             .SortKey = Column.Index - 1
             .SortOrder = lvwAscending
       Else
             If .SortOrder = lvwAscending Then
              .SortOrder = lvwDescending
             Else
              .SortOrder = lvwAscending
             End If
       End If
       .Sorted = -1
    End With
End Sub

Function pad(v, Optional l As Long = 8)
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < l Then
        pad = String(l - x, " ") & v
    Else
hell:
        pad = v
    End If
End Function

'supports %x, %c, %s, %d, %10d \t \n %%
Function printf(ByVal Msg As String, ParamArray vars()) As String

    Dim t
    Dim ret As String
    Dim i As Long
    
    Msg = Replace(Msg, Chr(0), Empty)
    Msg = Replace(Msg, "\t", vbTab)
    Msg = Replace(Msg, "\n", vbCrLf) 'simplified
    Msg = Replace(Msg, "%%", Chr(0))
    
    t = Split(Msg, "%")
    If UBound(t) <> UBound(vars) + 1 Then
        MsgBox "Format string mismatch.."
        Exit Function
    End If
    
    ret = t(0)
    For i = 1 To UBound(t)
        base = t(i)
        marker = ExtractSpecifier(base)
        If Len(marker) > 0 Then
            ret = ret & HandleMarker(base, marker, vars(i - 1))
        Else
            ret = ret & base
        End If
    Next
    
    ret = Replace(ret, Chr(0), "%")
    printf = ret
    
End Function

Private Function HandleMarker(base, ByVal marker, var) As String
    Dim newBase As String
    Dim mType As Integer
    Dim nVal As String
    Dim spacer As String
    Dim prefix As String
    Dim count As Long
    
    If Len(base) > Len(marker) Then
        newBase = Mid(base, Len(marker) + 1) 'remove the marker..
    End If
    
    mType = Asc(Mid(marker, Len(marker), 1))  'last character
    
    Select Case mType
        Case Asc("x"): nVal = Hex(var)
        Case Asc("X"): nVal = UCase(Hex(var))
        Case Asc("s"): nVal = var
        Case Asc("S"): nVal = UCase(var)
        Case Asc("c"): nVal = Chr(var)
        Case Asc("d"): nVal = var
        
        Case Else: nVal = var
    End Select
    
    If Len(marker) > 1 Then 'it has some more formatting involved..
        marker = Mid(marker, 1, Len(marker) - 1) 'trim off type
        If Left(marker, 1) = "0" Then
            spacer = "0"
            marker = Mid(marker, 2)
        Else
            spacer = " "
        End If
        count = CLng(marker) - Len(nVal)
        If count > 0 Then prefix = String(count, spacer)
    End If
    
    HandleMarker = prefix & nVal & newBase
            
End Function

Private Function ExtractSpecifier(v)
    
    Dim ret As String
    Dim b() As Byte
    Const LANG_US = &H409
    If Len(v) = 0 Then Exit Function
    
    b() = StrConv(v, vbFromUnicode, LANG_US)
    
    For i = 0 To UBound(b)
        ret = ret & Chr(b(i))
        If b(i) = Asc("x") Then Exit For
        If b(i) = Asc("X") Then Exit For
        If b(i) = Asc("c") Then Exit For
        If b(i) = Asc("s") Then Exit For
        If b(i) = Asc("S") Then Exit For
        If b(i) = Asc("d") Then Exit For
    Next
    
    ExtractSpecifier = ret
    
End Function

Public Function ODBCBuildConnectionString(dbServer As dbServers, dbName As String, Optional server As String, Optional Port = 3306, Optional user As String, Optional pass As String) As String
    Dim dbPath As String, baseString As String, blnInlineAuth As Boolean
    
    Select Case dbServer
        Case Access
            baseString = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb)};DBQ=____;"
        Case FileDsn
            baseString = "FILEDSN=____;"
        Case DSN
            baseString = "DSN=____;"
        Case dBase
            baseString = "Driver={Microsoft dBASE Driver (*.dbf)};DriverID=277;Dbq=____;"
        Case mysql
            baseString = "Driver={mySQL};Server=" & server & ";Port=" & Port & ";Stmt=;Option=16834;Database=____;"
        Case MsSql2k
            baseString = "Driver={SQL Server};Server=" & server & ";Database=____;"
        Case JetAccess2k
            baseString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=____;" & _
                         "User Id=" & user & ";" & _
                         "Password=" & pass & ";"
                         blnInlineAuth = True
    End Select
                         
        
    If Not blnInlineAuth Then
        If user <> Empty Then baseString = baseString & "Uid:" & user & ";"
        If pass <> Empty Then baseString = baseString & "Pwd:" & user & ";"
    End If
       
    '%AP% is like enviromental variable for app.path i am lazy :P
    dbPath = Replace(dbName, "%AP%", App.path)
    
    ODBCBuildConnectionString = Replace(baseString, "____", dbPath)
    
End Function


'should be dep safe..
Function CallCdecl(lpfn As Long, ParamArray args()) As Long
Attribute CallCdecl.VB_Description = "Call a CDecl DLL export. All arguments must be numeric. Convert strings to varptr(byte(0))"

    Dim asm() As String
    Dim stub() As Byte
    Dim i As Long
    Dim argSize As Byte
    Dim ret As Long
    Const PAGE_RWX      As Long = &H40
    Const MEM_COMMIT    As Long = &H1000
    Dim asmAddr As Long
    Dim sz As Long
    
    Const depSafe = True
    
    If lpfn = 0 Then Exit Function
    
    'push asm(), "CC"  'enable this to debug asm
    
    'we step through args backwards to preserve intutive ordering
    For i = UBound(args) To 0 Step -1
        If Not IsNumeric(args(i)) Then
            MsgBox "CallCdecl Invalid Parameter #" & i & " TypeName=" & TypeName(args(i))
            Exit Function
        End If
        push asm(), "68 " & lng2Hex(CLng(args(i)))  '68 90807000    PUSH 708090
        argSize = argSize + 4
    Next

    push asm(), "B8 " & lng2Hex(lpfn)        'B8 90807000    MOV EAX,708090
    push asm(), "FF D0"                      'FFD0           CALL EAX
    push asm(), "83 C4 " & Hex(argSize)      '83 C4 XX       add esp, XX     'cleanup args
    push asm(), "C2 10 00"                   'C2 10 00       retn 10h        'cleanup our callwindowproc args
    
    stub() = toBytes(Join(asm, " "))
    
    If Not depSafe Then
        CallCdecl = CallAsm(stub(0), 0, 0, 0, 0)
        Exit Function
    End If
    
    sz = UBound(stub) + 1
    asmAddr = VirtualAlloc(ByVal 0&, sz, MEM_COMMIT, PAGE_RWX)
    
    If asmAddr = 0 Then
        MsgBox "Failed to allocate RWE memory size: " & sz, vbInformation
        Exit Function
    End If
    
    RtlMoveMemory asmAddr, VarPtr(stub(0)), sz
    CallCdecl = CallAsmAddr(asmAddr, 0, 0, 0, 0)
    VirtualFree asmAddr, sz, 0
    
    
End Function

'endian swap and return spaced out hex string
Private Function lng2Hex(x As Long) As String
    Dim b(1 To 4) As Byte
    CopyMemory b(1), x, 4
    lng2Hex = Hex(b(1)) & " " & Hex(b(2)) & " " & Hex(b(3)) & " " & Hex(b(4))
End Function

'Function toBytes(x As String) As Byte()
'    Dim tmp() As String
'    Dim fx() As Byte
'    Dim i As Long
'
'    tmp = Split(x, " ")
'    ReDim fx(UBound(tmp))
'
'    For i = 0 To UBound(tmp)
'        fx(i) = CInt("&h" & tmp(i))
'    Next
'
'    toBytes = fx()
'
'End Function

Function toBytes(x As String) As Byte()
Attribute toBytes.VB_Description = "Accepts either spaced hex or straight hex strings, returns byte array"
    Dim tmp() As String
    Dim fx() As Byte
    Dim i As Long
    Dim y
    Dim sz As Long
    
    On Error Resume Next
    
    If InStr(x, " ") > 0 Then
        tmp = Split(x, " ")
        ReDim fx(UBound(tmp))
        
        For i = 0 To UBound(tmp)
            fx(i) = CInt("&h" & tmp(i))
        Next
    Else
        sz = Len(x) / 2
        ReDim fx(0 To sz - 1)
        y = Mid(x, 1, 2)
        fx(0) = CInt("&h" & y)
        For i = 1 To sz - 1
            y = Mid(x, (i * 2) + 1, 2)
            fx(i) = CInt("&h" & y)
        Next
    End If
    
    toBytes = fx()

End Function

Function isBinaryChar(b As Byte) As Boolean
     
    If b = 9 Or b = 10 Or b = 13 Then
        'tab, cr, lf are ok..
    ElseIf b >= &H20 And b <= &H7E Then
        'its a printable character and is ok..
    Else
        isBinaryChar = True
    End If
 
End Function
