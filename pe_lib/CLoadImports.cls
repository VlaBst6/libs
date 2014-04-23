VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLoadImports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com

Public Modules As New Collection  'of CImport
Private fhandle As Long 'my FileHandle
Private mParent As CPEEditor

Friend Sub LoadImports(fhand As Long, rvaImportDir As Long, parent As CPEEditor)
    On Error Resume Next
    
    Dim fOffsetIDT As Long
       
    Set mParent = parent
    
    fOffsetIDT = mParent.RvaToOffset(rvaImportDir)
    
    If fOffsetIDT < 1 Then
        MsgBox "Error getting file offset from rva? rva=" & rvaImportDir & " offset=" & fOffsetIDT, vbExclamation
        Exit Sub
    End If
    
    Dim idt As IMAGE_IMPORT_DIRECTORY
    
    fhandle = fhand
        
    fOffsetIDT = fOffsetIDT + 1 'cause 0 based file offset
    
    Get fhandle, fOffsetIDT, idt
    
    While idt.pFuncAry > 0
        LoadModuleImports idt
        fOffsetIDT = fOffsetIDT + Len(idt)
        'Debug.Print "offset=" & Hex(fOffsetIDT)
        Get fhandle, fOffsetIDT, idt
    Wend
    
End Sub

Private Sub LoadModuleImports(iid As IMAGE_IMPORT_DIRECTORY)
    'On Error GoTo hell
    
    Dim m  As New CImport

    With iid
        'Debug.Print "pFuncAry=" & Hex(.pFuncAry)
        m.pLookupTable = mParent.RvaToOffset(.pFuncAry)
        m.DllName = ScanString(mParent.RvaToOffset(.pDllName))
        m.FirstThunk = .pThunk
    End With
    
    Dim ImportLookup As Long, i As Long
    Dim isOrdial As Boolean, ImportFx As String
    
    i = m.pLookupTable + 1
    
    Get fhandle, i, ImportLookup
    isOrdial = CBool(ImportLookup And &H80000000) 'bitmask for high order bit
    ImportLookup = ImportLookup And &H7FFFFFFF    'rest of the bits (either ordial or rva to string name)

        
    While ImportLookup > 0 Or isOrdial
    
        If isOrdial Then
            ImportFx = "@" & Hex(ImportLookup)
            'Debug.Print ImportFx
            m.functions.Add ImportFx
        Else
            ImportFx = ScanString(mParent.RvaToOffset(ImportLookup), True)
            'Debug.Print ImportFx
            m.functions.Add ImportFx
        End If
        
        i = i + 4 'next record is 4 bytes away
        Get fhandle, i, ImportLookup
        isOrdial = CBool(ImportLookup And &H80000000) 'bitmask for high order bit
        ImportLookup = ImportLookup And &H7FFFFFFF    'rest of the bits (either ordial or rva to string name)

    Wend
    
    Modules.Add m, CStr(m.DllName)

Exit Sub
hell: MsgBox Err.Description

End Sub

Private Function ScanString(offset As Long, Optional getHint As Boolean = False) As String
    Dim b As Byte, i As Long, ret As String
    Dim Hint As Integer '2 bytes
     
    i = offset + 1
    
    If getHint Then
        Get fhandle, i, Hint
        i = i + 2
    End If
    
    Get fhandle, i, b
     
    While b <> 0
        ret = ret & Chr(b)
        Get fhandle, , b
    Wend
    
    ScanString = ret
 
End Function


Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
