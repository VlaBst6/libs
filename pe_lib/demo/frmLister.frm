VERSION 5.00
Begin VB.Form frmLister 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   5895
   End
End
Attribute VB_Name = "frmLister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ShowList(ary() As String)
    
    Text1 = Join(ary, vbCrLf)
    Me.Show 1
    
End Sub
