VERSION 5.00
Begin VB.Form frmlist 
   BackColor       =   &H000040C0&
   Caption         =   "List Of Games"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   13680
   LinkTopic       =   "Form5"
   Picture         =   "list.frx":0000
   ScaleHeight     =   2078.479
   ScaleMode       =   0  'User
   ScaleWidth      =   246.99
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   5520
      Picture         =   "list.frx":25DFB
      ScaleHeight     =   1065
      ScaleWidth      =   4545
      TabIndex        =   6
      Top             =   360
      Width           =   4575
   End
   Begin VB.OptionButton opt2 
      BackColor       =   &H000000C0&
      Caption         =   "  Puzzle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6000
      Picture         =   "list.frx":27AAC
      TabIndex        =   5
      Top             =   8520
      Width           =   1575
   End
   Begin VB.OptionButton opt1 
      BackColor       =   &H0000FFFF&
      Caption         =   "  Tic Tac Toe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      Picture         =   "list.frx":28165
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12720
      Picture         =   "list.frx":2881E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   4215
      Left            =   8640
      Picture         =   "list.frx":28E72
      ScaleHeight     =   4155
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   5160
      Width           =   6735
   End
   Begin VB.PictureBox Picture1 
      Height          =   4935
      Left            =   600
      Picture         =   "list.frx":3585C
      ScaleHeight     =   4875
      ScaleWidth      =   4995
      TabIndex        =   1
      Top             =   2160
      Width           =   5055
   End
   Begin VB.CommandButton cmdback1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Picture         =   "list.frx":3DC86
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdback1_Click()
frmintro.Show
Me.Hide
End Sub

Private Sub cmdnext_Click()

If opt1.Value Then
  Me.Hide
  frmsub1.Show
ElseIf opt2.Value Then
  Me.Hide
  frmsub2.Show
Else
  MsgBox (" Please Select One Option.")
  
End If

End Sub


