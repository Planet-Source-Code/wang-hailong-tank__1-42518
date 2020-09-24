VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Height          =   3330
      Left            =   0
      Pattern         =   "*.txt"
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim mstr As String
mstr = File1.filename
If mstr <> "" Then Form1.readit mstr
Unload Form2

End Sub

Private Sub Command2_Click()
Unload Form2
End Sub

Private Sub File1_DblClick()
Command1_Click
End Sub

Private Sub Form_Load()
File1.Path = "map"
End Sub
