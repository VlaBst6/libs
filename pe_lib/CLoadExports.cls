VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLoadExports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com

Public functions As New Collection  'of CExport
Private mParent As CPEEditor

Friend Sub LoadExports(fhandle As Long, ByVal rvaExportDirectoryTable As Long, parent As CPEEditor)
    
    Dim offsetExportTable As Long
    
    Set mParent = parent
    
    If rvaExportDirectoryTable < 1 Then Exit Sub

    Dim ied As IMAGE_EXPORT_DIRECTORY
    Dim firstOrdial As Long
    Dim offset As Long
    Dim pointer As Long
    Dim i As Long
    
    Dim functions() As Long
    Dim names() As String
    Dim ordials() As Integer '2bytes each
    
    offsetExportTable = mParent.RvaToOffset(rvaExportDirectoryTable) + 1
    
    Get fhandle, offsetExportTable, ied
    firstOrdial = ied.base
    offset = mParent.RvaToOffset(ied.AddressOfFunctions) + 1
        
    For i = 0 To ied.NumberOfFunctions - 1
        Get fhandle, (offset + (i * 4)), pointer
        push functions(), pointer
        'Debug.Print "Fx Addr " & i & ":" & Hex(pointer)
    Next

    offset = mParent.RvaToOffset(ied.AddressOfNames) + 1
    
    For i = 0 To ied.NumberOfNames - 1
        Get fhandle, (offset + (i * 4)), pointer
        push names(), ExtractExportName(fhandle, mParent.RvaToOffset(pointer))
    Next
        
    If ied.NumberOfFunctions > 0 Then
        ReDim ordials(ied.NumberOfFunctions - 1)
        Get fhandle, mParent.RvaToOffset(ied.AddressOfNameOrdinals) + 1, ordials
            
        For i = 0 To ied.NumberOfFunctions - 1
           AddExport functions(i), names(i), (ordials(i) + firstOrdial)
        Next
    End If
    
    
End Sub

Private Function ExtractExportName(f As Long, pointer As Long) As String
     Dim b As Byte
     Dim tmp As String
     
        If pointer <= 0 Or pointer >= LOF(f) Then
            ExtractExportName = "[No Name]"
        Else
            Get f, (pointer + 1), b
            While b <> 0
                tmp = tmp & Chr(b)
                Get f, , b
            Wend
            ExtractExportName = tmp
        End If
        
End Function


Private Sub AddExport(address As Long, name As String, Optional ordial As Long)
    Dim m  As New CExport
    m.FunctionAddress = address
    m.FunctionName = name
    m.FunctionOrdial = ordial
    functions.Add m
End Sub



Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub
