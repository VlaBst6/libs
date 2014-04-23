Attribute VB_Name = "modPEStructs"
Option Explicit
'these struct definitions were taken from VB debugger sample by VF-fCRO
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=42422&lngWId=1

Public Const IMAGE_NUMBEROF_DIRECTORY_ENTRIES = 16
Public Const IMAGE_SIZEOF_SHORT_NAME = 8
Public Const IMAGE_NT_OPTIONAL_HDR32_MAGIC = &H10B

Public Type IMAGEDOSHEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_ovno As Integer
    e_res(1 To 4) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(1 To 10)    As Integer
    e_lfanew As Long
End Type

Public Type IMAGE_EXPORT_DIRECTORY
    Characteristics As Long
    TimeDateStamp As Long
    MajorVersion As Integer
    MinorVersion As Integer
    name As Long
    base As Long
    NumberOfFunctions As Long
    NumberOfNames As Long
    AddressOfFunctions As Long
    AddressOfNames As Long
    AddressOfNameOrdinals As Long
End Type

Public Type IMAGE_IMPORT_DIRECTORY
    pFuncAry As Long
    timestamp As Long
    forwarder As Long
    pDllName As Long
    pThunk As Long
End Type

Public Type IMAGE_SECTION_HEADER
    nameSec As String * 6
    PhisicalAddress As Integer
    VirtualSize As Long
    VirtualAddress As Long
    SizeOfRawData As Long
    PointerToRawData As Long
    PointerToRelocations As Long
    PointerToLinenumbers As Long
    NumberOfRelocations As Integer
    NumberOfLinenumbers As Integer
    Characteristics As Long
End Type

Public Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDateStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Public Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    size As Long
End Type


Public Type IMAGE_OPTIONAL_HEADER
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitializedData As Long
    SizeOfUninitializedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Win32VersionValue As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaderFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY
End Type

'Enum eDATA_DIRECTORY
'    Export_Table = 0
'    Import_Table = 1
'    Resource_Table = 2
'    Exception_Table = 3
'    Certificate_Table = 4
'    Relocation_Table = 5
'    Debug_Data = 6
'    Architecture_Data = 7
'    Machine_Value = 8        '(MIPS_GP)
'    TLS_Table = 9
'    Load_Configuration_Table = 10
'    Bound_Import_Table = 11
'    Import_Address_Table = 12
'    Delay_Import_Descriptor = 13
'    COM_Runtime_Header = 14
'    Reserved = 15
'End Enum


Public Type IMAGE_NT_HEADERS
    Signature As String * 4
    FileHeader As IMAGE_FILE_HEADER
    OptionalHeader As IMAGE_OPTIONAL_HEADER
End Type



Function toHex(ParamArray elems())
    On Error Resume Next
    Dim i As Long
    For i = 0 To UBound(elems)
        elems(i).Text = Hex(elems(i).Text)
    Next
End Function



Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function GetHextxt(t As TextBox, v As Long) As Boolean
    
    On Error Resume Next
    v = CLng("&h" & t)
    If Err.Number > 0 Then
        MsgBox "Error " & t.Text & " is not valid hex number", vbInformation
        Exit Function
    End If
    
    GetHextxt = True
    
End Function



Sub Enable(t As TextBox, Optional enabled = True)
    t.BackColor = IIf(enabled, vbWhite, &H80000004)
    t.enabled = enabled
    t.Text = Empty
End Sub

Function Align(ByVal valu) As Long
    While valu Mod 16 <> 0
        valu = valu + 1
    Wend
    Align = valu
End Function

Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function


Sub ConfigureListView(lv As Object)
        
        Dim i As Integer
        
        lv.FullRowSelect = True
        lv.GridLines = True
        lv.HideColumnHeaders = False
        lv.View = 3 'lvwReport
    
        lv.ColumnHeaders.Clear
        lv.ColumnHeaders.Add , , "Section Name"
        lv.ColumnHeaders.Add , , "VirtualAddr"
        lv.ColumnHeaders.Add , , "VirtualSize"
        lv.ColumnHeaders.Add , , "RawOffset"
        lv.ColumnHeaders.Add , , "RawSize"
        lv.ColumnHeaders.Add , , "Characteristics"
        
        lv.Width = (1250 * 6) + 250
        lv.Height = 1800
        
        For i = 1 To 6
            lv.ColumnHeaders(i).Width = 1250
        Next
        
End Sub

Sub FilloutListView(lv As Object, Sections As Collection)
        
    If Sections.Count = 0 Then
        MsgBox "Sections not loaded yet"
        Exit Sub
    End If
    
    Dim cs As CSection, li As Object 'ListItem
    lv.ListItems.Clear
    
    For Each cs In Sections
        Set li = lv.ListItems.Add(, , cs.nameSec)
        li.SubItems(1) = Hex(cs.VirtualAddress)
        li.SubItems(2) = Hex(cs.VirtualSize)
        li.SubItems(3) = Hex(cs.PointerToRawData)
        li.SubItems(4) = Hex(cs.SizeOfRawData)
        li.SubItems(5) = Hex(cs.Characteristics)
    Next
    
    Dim i As Integer
    For i = 1 To lv.ColumnHeaders.Count
        lv.ColumnHeaders(i).Width = 1000
    Next
    With lv.ColumnHeaders(i - 1)
        .Width = lv.Width - .Left - 100
    End With
    
    
End Sub

