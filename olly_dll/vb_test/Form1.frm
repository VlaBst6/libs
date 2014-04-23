VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5580
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   1515
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1620
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Height          =   1515
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'dzzie@yahoo.com
'http://sandsprite.com


Private Type t_asmmodel
    code(15) As Byte
    mask(15) As Byte
    Length As Long
    jmpsize As Long
    jmpoffset As Long
    jmppos As Long
End Type

Private Type t_Disasm  '         // Results of disassembling
  ip As Long  '                  // Instrucion pointer
  dump As String * 256        '  // Hexadecimal dump of the command
  result As String * 256
  comment As String * 256     '  // Brief comment
  cmdtype As Long '              // One of C_xxx
  memtype As Long '              // Type of addressed variable in memory
  nprefix As Long '              // Number of prefixes
  indexed As Long '              // Address contains register(s)
  jmpconst As Long '             // Constant jump address
  jmptable As Long '             // Possible address of switch table
  adrconst As Long '             // Constant part of address
  immconst As Long '             // Immediate constant
  zeroconst As Long '            // Whether contains zero constant
  fixupoffset As Long '          // Possible offset of 32-bit fixups
  fixupsize As Long '            // Possible total size of fixups or 0
  error As Long '                // Error while disassembling command
  warnings As Long '             // Combination of DAW_xxx
End Type

Enum disasmMode
    DISASM_SIZE = 0     '            // Determine command size only
    DISASM_DATA = 1     '            // Determine size and analysis data
    DISASM_FILE = 3     '            // Disassembly, no symbols
    DISASM_CODE = 4     '            // Full disassembly
End Enum


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function disasm Lib "olly.dll" Alias "Disasm" ( _
        ByRef src As Byte, ByVal srcsize As Long, ByVal ip As Long, _
        disasm As t_Disasm, dMode As disasmMode) As Long

Private Declare Sub VB_SetOptions Lib "olly.dll" ( _
        Optional ByVal isideal As Long = 0, Optional ByVal isLower As Long = 0, _
        Optional doTabs As Long = 0, Optional ByVal dispseg As Long = 1)

Private Declare Function Assemble Lib "olly.dll" ( _
        ByVal CMD As String, ByVal ip As Long, model As t_asmmodel, _
        ByVal attempt As Long, ByVal constsize As Long, ByVal errtext As String) As Long


Const base As Long = &H400000

Private Sub Command1_Click()
    AsmBlock Text1
End Sub

Private Sub Form_Load()

     VB_SetOptions
     
     Dim test() As String
     push test, "MOV EAX, [EAX]"
     push test, "CALL 4012bb"
     push test, "MOV AL, CL"
     push test, "JNZ 401567"
     
     Text1 = Join(test, vbCrLf)
     
     Command1_Click
     
End Sub

Sub AsmBlock(inst As String)
    Dim tmp() As String
    Dim i As Integer
    Dim b() As Byte
    Dim offset As Long
    Dim leng As Long
    Dim errMsg As String
    Dim dis() As String
    
    On Error GoTo hell
    
    offset = base
    tmp = Split(inst, vbCrLf)
    
    For i = 0 To UBound(tmp)
        b() = AsmLine(tmp(i), offset, leng, errMsg)
        push dis(), Hex(offset) & "    " & dump(b) & vbTab & DisasmBytes(b, offset)
        offset = offset + leng
        Me.Caption = errMsg
    Next
    
    Text2 = Join(dis, vbCrLf)
    
hell:
End Sub

Function AsmLine(line As String, offset As Long, asmLen As Long, errMsg As String) As Byte()
    Dim leng As Long
    Dim am As t_asmmodel
    Dim b() As Byte
    Dim i As Integer
    
    errMsg = String(255, " ")
    asmLen = Assemble(line, offset, am, 0, 0, errMsg)
    
    
    
    If asmLen < 0 Then Exit Function
    
    ReDim b(1 To asmLen)
    CopyMemory b(1), am.code(0), asmLen
    
    AsmLine = b()
    
End Function

 
 Function DisasmBytes(b() As Byte, offset) As String
    Dim da As t_Disasm
    Dim x As Long
    Dim src As String
     
    x = disasm(b(1), UBound(b), offset, da, DISASM_CODE)
    DisasmBytes = da.result
    
    x = InStr(DisasmBytes, Chr(0)) - 1
    DisasmBytes = Mid(DisasmBytes, 1, x)
    
 End Function


Function dump(b() As Byte)
    Dim tmp As String
    Dim i As Integer
    Dim x As String
    
    For i = LBound(b) To UBound(b)
        x = Hex(b(i))
        If Len(x) = 1 Then x = "0" & x
        tmp = tmp & x & " "
    Next
    
    i = Len(tmp)
    
    If i < 30 Then tmp = tmp & Space(30 - i)
    
    dump = tmp
    
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



Private Sub Text1_LostFocus()
    Command1_Click
End Sub
