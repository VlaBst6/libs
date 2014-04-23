VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPEEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com
'
'These pe editor classes are v 0.21

Public OptionalHeader As COptionalHeader
Public Imports As CLoadImports
Public Exports As CLoadExports
Public Sections As Collection

Public CanAddSection As Boolean
Public errMessage As String

Private DOSHEADER As IMAGEDOSHEADER
Private NTHEADER As IMAGE_NT_HEADERS

Private myFile As String

Sub ShowOffsetCalculator(Optional modal As Boolean = False)
    frmOffsets.Initilize Me, modal
End Sub

Property Get CompiledDate() As String

    On Error Resume Next
    Dim base As Date
    Dim compiled As Date
    
    base = DateSerial(1970, 1, 1)
    compiled = DateAdd("s", NTHEADER.FileHeader.TimeDateStamp, base)
    CompiledDate = "GMT: " & Format(compiled, "ddd mmm d h:nn:ss yyyy")

End Property

Function LoadFile(fpath As String) As Boolean
    On Error GoTo hell
        
        Dim i As Long
        Dim f As Long
        Dim udtSH As IMAGE_SECTION_HEADER
        
        errMessage = Empty
        If Not FileExists(fpath) Then Err.Raise 1
            
        Set OptionalHeader = New COptionalHeader
        Set Imports = New CLoadImports
        Set Exports = New CLoadExports
        Set Sections = New Collection

        f = FreeFile
        
        Open fpath For Binary Access Read As f
        Get f, , DOSHEADER
        Get f, DOSHEADER.e_lfanew + 1, NTHEADER
        
        'If NTHEADER.Signature <> "PE" & Chr(0) & Chr(0) Then Err.Raise 2
        
        OptionalHeader.LoadStruc NTHEADER.OptionalHeader
        LoadSections f, NTHEADER
        Imports.LoadImports f, OptionalHeader.ddVirtualAddress(Import_Table), Me
        Exports.LoadExports f, OptionalHeader.ddVirtualAddress(Export_Table), Me
        
        Close f
        LoadFile = True
        myFile = fpath
        
Exit Function
hell:
    
    Close f
    
    Select Case Err.Number
        Case 1: errMessage = "File not found: " & fpath
        Case 2: errMessage = "Invalid PE signature: " & NTHEADER.Signature
        Case Else: errMessage = Err.Description
    End Select
    
End Function

Function SaveChanges() As Boolean
   On Error GoTo hell

    errMessage = Empty
    If Len(myFile) = 0 Then Err.Raise 1
    If Not FileExists(myFile) Then Err.Raise 2

    Dim i As Long
    Dim f As Long
    Dim udtSH As IMAGE_SECTION_HEADER
    Dim csh As CSection

    f = FreeFile

    NTHEADER.OptionalHeader = OptionalHeader.GetStruc

    Open myFile For Binary As f
    Put f, , DOSHEADER
    Put f, DOSHEADER.e_lfanew + 1, NTHEADER

    For i = 1 To NTHEADER.FileHeader.NumberOfSections
        Set csh = Sections.Sections(i)
        udtSH = csh.GetStruc
        Put f, , udtSH
    Next

    SaveChanges = True

Exit Function
hell:

    Close f
    
    Select Case Err.Number
        Case 1: errMessage = "No file is currently Loaded"
        Case 2: errMessage = "File not found: " & myFile
        Case Else: errMessage = Err.Description
    End Select

End Function



'Sub AddSection(csh As CSection)  'not tested yet
'    Dim cur As Integer'
'
 '   cur = NTHEADER.FileHeader.NumberOfSections
'    cur = cur + 1'
'
'    SectionHeaders.Add csh, "sect:" & cur
'    NTHEADER.FileHeader.NumberOfSections = cur'
'
'End Sub

Private Sub LoadSections(fhand As Long, nth As IMAGE_NT_HEADERS)
    
    Dim i As Long
    Dim udtSH As IMAGE_SECTION_HEADER
    Dim CSect As CSection
    
    Set Sections = New Collection
    
    For i = 1 To nth.FileHeader.NumberOfSections
        Get fhand, , udtSH
        Set CSect = New CSection
        CSect.LoadStruc udtSH
        CSect.index = i
        Sections.Add CSect
    Next
    
    CanAddSection = False
    
    Dim b() As Byte
    ReDim b(Len(udtSH))
    Get fhand, , b() ' buffer where next section would sit
    
    'if all empty then is room to add new section
    For i = 0 To UBound(b)
        If b(i) <> 0 Then Exit Sub
    Next
    
    CanAddSection = True
    
End Sub

 
 
Function RvaToOffset(ByVal rva As Long, Optional ByRef VirtualSectionBase As Long, Optional sectName As String)
        
    If Sections.Count = 0 Then
        MsgBox "No Sections Loaded", vbInformation
        Exit Function
    End If
    
    Dim c As CSection
    Dim vBase As Long, vSize As Long, vMax As Long
    
    For Each c In Sections
       vBase = c.VirtualAddress
       vSize = c.VirtualSize
       vMax = vBase + vSize
       
       If rva >= vBase And rva < vMax Then 'right section
            rva = rva - vBase
            rva = rva + c.PointerToRawData
            RvaToOffset = rva
            VirtualSectionBase = vBase
            sectName = c.nameSec
            Exit Function
       End If
    Next
       
    
End Function

Function OffsetToRVA(ByVal fOffset As Long, Optional sectName As String)
    
    If Sections.Count = 0 Then
        MsgBox "No Sections Loaded", vbInformation
        Exit Function
    End If
    
    Dim c As CSection
    Dim rBase As Long, rSize As Long, rMax As Long
    
    For Each c In Sections
       rBase = c.PointerToRawData
       rSize = c.SizeOfRawData
       rMax = rBase + rSize
       
       If fOffset >= rBase And fOffset < rMax Then 'right section
            OffsetToRVA = c.VirtualAddress + fOffset - rBase
            sectName = c.nameSec
            Exit Function
       End If
    Next

End Function

Function GetSectionFromRVA(rva As Long) As CSection
      
    If Sections.Count = 0 Then
        MsgBox "No Sections Loaded", vbInformation
        Exit Function
    End If
    
    Dim c As CSection
    Dim vBase As Long, vSize As Long, vMax As Long
    
    For Each c In Sections
       vBase = c.VirtualAddress
       vSize = c.VirtualSize
       vMax = vBase + vSize
       
       If rva >= vBase And rva < vMax Then 'right section
            Set GetSectionFromRVA = c
            Exit Function
       End If
    Next
    
End Function

Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function




