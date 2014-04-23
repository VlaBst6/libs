VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lv 
      Height          =   2955
      Left            =   60
      TabIndex        =   3
      Top             =   2580
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Offset"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Dump"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Disasm"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   5580
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Assemble"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   5580
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2535
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

Const base As Long = &H400000

Dim asm As New CAssembler
Dim dsm As New CDisassembler

'assembler supports:
'  call $+x style labels for forward (or back) jumps/calls from cur pos
'  jmps or calls backwards to labels
'  can optionally add other address labels such as for api calls or from dbg info

'assembler and disassembler functionality provided from
'GPL ollydbg disasm sources Copyright (C) 2001 Oleh Yuschuk.
'
'These are free sources that are part of his freeware debugger
'Unaltered copies can be downloaded here:
'
'http://home.t-online.de/home/Ollydbg/


Private Sub Form_Load()

   asm.AddApi "MyTestAPI", &HDEADBEEF
   
End Sub

Private Sub Command1_Click()
    Dim op() As Byte
    
    Dim disasm As Collection
    Dim ci As CInstruction
    Dim li As ListItem
    
    lv.ListItems.Clear
     
    op() = asm.AsmBlock(Text1, base)
    
    Set disasm = dsm.DisasmBlock(op, base)
    For Each ci In disasm
        Set li = lv.ListItems.Add(, , ci.offset)
        li.SubItems(1) = ci.dump
        li.SubItems(2) = ci.command
    Next
    
    Text3 = dumpIt(op)
    
End Sub
 
Private Sub Text1_LostFocus()
    Command1_Click
End Sub

Function dumpIt(b() As Byte)
    Dim tmp As String
    Dim i As Integer
    Dim x As String
    
    On Error GoTo hell
    
    For i = LBound(b) To UBound(b)
        x = Hex(b(i))
        If Len(x) = 1 Then x = "0" & x
        tmp = tmp & x & " "
    Next
    
    i = Len(tmp)
    
    'If i < 30 Then tmp = tmp & Space(30 - i)
    
    dumpIt = tmp
hell:

End Function



