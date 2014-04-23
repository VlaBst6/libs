VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COptionalHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com

Public Magic As Integer
Public MajorLinkerVersion As Byte
Public MinorLinkerVersion As Byte
Public SizeOfCode As Long
Public SizeOfInitializedData As Long
Public SizeOfUninitializedData As Long
Public AddressOfEntryPoint As Long
Public BaseOfCode As Long
Public BaseOfData As Long
Public ImageBase As Long
Public SectionAlignment As Long
Public FileAlignment As Long
Public MajorOperatingSystemVersion As Integer
Public MinorOperatingSystemVersion As Integer
Public MajorImageVersion As Integer
Public MinorImageVersion As Integer
Public MajorSubsystemVersion As Integer
Public MinorSubsystemVersion As Integer
Public Win32VersionValue As Long
Public SizeOfImage As Long
Public SizeOfHeaders As Long
Public CheckSum As Long
Public Subsystem As Integer
Public DllCharacteristics As Integer
Public SizeOfStackReserve As Long
Public SizeOfStackCommit As Long
Public SizeOfHeapReserve As Long
Public SizeOfHeapCommit As Long
Public LoaderFlags As Long
Public NumberOfRvaAndSizes As Long
Private DataDirectory(0 To 15) As IMAGE_DATA_DIRECTORY

Public Enum eDATA_DIRECTORY
    Export_Table = 0
    Import_Table = 1
    Resource_Table = 2
    Exception_Table = 3
    Certificate_Table = 4
    Relocation_Table = 5
    Debug_Data = 6
    Architecture_Data = 7
    Machine_Value = 8        '(MIPS_GP)
    TLS_Table = 9
    Load_Configuration_Table = 10
    Bound_Import_Table = 11
    Import_Address_Table = 12
    Delay_Import_Descriptor = 13
    COM_Runtime_Header = 14
    Reserved = 15
End Enum

Friend Function LoadStruc(oh As IMAGE_OPTIONAL_HEADER)
    Dim i As Integer
    
    With oh
        Magic = .Magic
        MajorLinkerVersion = .MajorLinkerVersion
        MinorLinkerVersion = .MinorLinkerVersion
        SizeOfCode = .SizeOfCode
        SizeOfInitializedData = .SizeOfInitializedData
        SizeOfUninitializedData = .SizeOfUninitializedData
        AddressOfEntryPoint = .AddressOfEntryPoint
        BaseOfCode = .BaseOfCode
        BaseOfData = .BaseOfData
        ImageBase = .ImageBase
        SectionAlignment = .SectionAlignment
        FileAlignment = .FileAlignment
        MajorOperatingSystemVersion = .MajorOperatingSystemVersion
        MinorOperatingSystemVersion = .MinorOperatingSystemVersion
        MajorImageVersion = .MajorImageVersion
        MinorImageVersion = .MinorImageVersion
        MajorSubsystemVersion = .MajorSubsystemVersion
        MinorSubsystemVersion = .MinorSubsystemVersion
        Win32VersionValue = .Win32VersionValue
        SizeOfImage = .SizeOfImage
        SizeOfHeaders = .SizeOfHeaders
        CheckSum = .CheckSum
        Subsystem = .Subsystem
        DllCharacteristics = .DllCharacteristics
        SizeOfStackReserve = .SizeOfStackReserve
        SizeOfStackCommit = .SizeOfStackCommit
        SizeOfHeapReserve = .SizeOfHeapReserve
        SizeOfHeapCommit = .SizeOfHeapCommit
        LoaderFlags = .LoaderFlags
        NumberOfRvaAndSizes = .NumberOfRvaAndSizes
        For i = 0 To 15
            DataDirectory(i) = .DataDirectory(i)
        Next
   End With
    
    
End Function

Friend Function GetStruc() As IMAGE_OPTIONAL_HEADER
    Dim i As Integer
    
    With GetStruc
        .Magic = Magic
        .MajorLinkerVersion = MajorLinkerVersion
        .MinorLinkerVersion = MinorLinkerVersion
        .SizeOfCode = SizeOfCode
        .SizeOfInitializedData = SizeOfInitializedData
        .SizeOfUninitializedData = SizeOfUninitializedData
        .AddressOfEntryPoint = AddressOfEntryPoint
        .BaseOfCode = BaseOfCode
        .BaseOfData = BaseOfData
        .ImageBase = ImageBase
        .SectionAlignment = SectionAlignment
        .FileAlignment = FileAlignment
        .MajorOperatingSystemVersion = MajorOperatingSystemVersion
        .MinorOperatingSystemVersion = MinorOperatingSystemVersion
        .MajorImageVersion = MajorImageVersion
        .MinorImageVersion = MinorImageVersion
        .MajorSubsystemVersion = MajorSubsystemVersion
        .MinorSubsystemVersion = MinorSubsystemVersion
        .Win32VersionValue = Win32VersionValue
        .SizeOfImage = SizeOfImage
        .SizeOfHeaders = SizeOfHeaders
        .CheckSum = CheckSum
        .Subsystem = Subsystem
        .DllCharacteristics = DllCharacteristics
        .SizeOfStackReserve = SizeOfStackReserve
        .SizeOfStackCommit = SizeOfStackCommit
        .SizeOfHeapReserve = SizeOfHeapReserve
        .SizeOfHeapCommit = SizeOfHeapCommit
        .LoaderFlags = LoaderFlags
        .NumberOfRvaAndSizes = NumberOfRvaAndSizes
         For i = 0 To 15
            .DataDirectory(i) = DataDirectory(i)
         Next
    End With

End Function

Property Get ddVirtualAddress(i As eDATA_DIRECTORY) As Long
    ddVirtualAddress = DataDirectory(i).VirtualAddress
End Property

Property Get ddSize(i As eDATA_DIRECTORY) As Long
    ddSize = DataDirectory(i).size
End Property

