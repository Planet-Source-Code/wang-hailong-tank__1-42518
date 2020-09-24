VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   8625
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command5 
      Caption         =   "new"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "save"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "open"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "¡ú"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¡û"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   5415
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5355
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.Image Image2 
         Height          =   390
         Left            =   2760
         Picture         =   "Form1.frx":0000
         Top             =   4920
         Width           =   450
      End
   End
   Begin VB.Label Label1 
      Caption         =   "no map was open"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   5
      Left            =   1680
      Picture         =   "Form1.frx":099A
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   4
      Left            =   1320
      Picture         =   "Form1.frx":0D1C
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   3
      Left            =   960
      Picture         =   "Form1.frx":109E
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   2
      Left            =   600
      Picture         =   "Form1.frx":1420
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   240
      Picture         =   "Form1.frx":176E
      Top             =   5640
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   120
      Top             =   5520
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type maptype
    X As Integer
    Y As Integer
    shape As Long
End Type
Dim map(1 To 598) As maptype
Dim w As Long
Dim h As Long
Dim shape As Long
Dim filename As String
Private Sub Command1_Click()
Dim i As Long
Picture1.Cls
Dim s As String
s = ""
Dim s2 As String

For i = 1 To 598
    map(i).shape = 0
Next

s = Text1.Text
For i = 1 To Len(s)
    map(i).shape = Mid(s, i, 1)
    If map(i).shape <> 0 Then
        Picture1.PaintPicture Image1(map(i).shape), map(i).X, map(i).Y
    End If
Next i

shape = 1
For i = 1 To Image1.UBound
    Image1(i).BorderStyle = 0
Next i
Image1(1).BorderStyle = 1
End Sub

Private Sub Command2_Click()
Dim s As String
Dim i As Long
For i = 1 To 598
    s = s & map(i).shape
Next i
Text1.Text = s
End Sub

Private Sub Command3_Click()
Form2.Show

End Sub
Sub readit(mstr As String)
Dim maps As String
Dim c As Long
If mstr <> "" Then
    filename = mstr
    c = Val(Mid(mstr, 4, Len(mstr) - 7))
    Open "map\" & mstr For Input As #1
        Input #1, maps
    Close
    Label1.Caption = "in editing the round" & c
    Text1.Text = maps
    Command1_Click
End If
End Sub

Private Sub Command4_Click()
Command2_Click
Dim maps As String
Dim i As Long
maps = Text1.Text
If filename <> "" Then
    If Dir("map\" & filename) <> "" Then Kill "map\" & filename
    Open "map\" & filename For Output As #1
        Print #1, maps
    Close
End If
End Sub

Private Sub Command5_Click()
Dim i As Long
Dim j As Long

Dim mstr As String
filename = ""
Text1.Text = ""
Picture1.Picture = LoadPicture("")
Picture1.Cls
For i = 1 To 598
    map(i).shape = 0
Next i
For j = 1 To 1000
    mstr = "map" & j & ".txt"
    If Dir("map\" & mstr) = "" Then
        filename = mstr
        Label1.Caption = "in editing the round" & j
        Exit For
    End If
Next j
End Sub

Private Sub Form_Load()
Dim i As Long
For i = 1 To 598
    w = Picture1.Width / 26
    h = Picture1.Height / 23
    map(i).X = ((i - 1) Mod 26) * w
    map(i).Y = ((i - 1) \ 26) * h
'Picture1.PaintPicture Image1(0).Picture, map(i).X, map(i).Y, w, h, 0, 0, w, h
Next i
Image1(1).BorderStyle = 1
shape = 1
End Sub

Private Sub Image1_Click(Index As Integer)
Dim i As Long
For i = 1 To Image1.UBound
    Image1(i).BorderStyle = 0
Next i
Image1(Index).BorderStyle = 1
shape = Index
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
For i = 1 To 598
    If X > map(i).X And X <= map(i).X + w Then
        If Y > map(i).Y And Y <= map(i).Y + h Then
            If Button = 1 Then
                Picture1.PaintPicture Image1(shape).Picture, map(i).X, map(i).Y, w, h
                map(i).shape = shape
            ElseIf Button = 2 Then
                Picture1.Line (map(i).X, map(i).Y)-(map(i).X + w, map(i).Y + h), , BF
                map(i).shape = 0
            End If
        End If
    End If
Next i

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
For i = 1 To 598
    If X > map(i).X And X <= map(i).X + w Then
        If Y > map(i).Y And Y <= map(i).Y + h Then
            If Button = 1 Then
                Picture1.PaintPicture Image1(shape).Picture, map(i).X, map(i).Y, w, h
                map(i).shape = shape
            ElseIf Button = 2 Then
                Picture1.Line (map(i).X, map(i).Y)-(map(i).X + w, map(i).Y + h), , BF
                map(i).shape = 0
            End If
        End If
    End If
Next i

End Sub
