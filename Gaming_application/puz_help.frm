VERSION 5.00
Begin VB.Form frmpuzhelp 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Puzzle  Help"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13875
   FillStyle       =   2  'Horizontal Line
   LinkTopic       =   "Form9"
   Picture         =   "puz_help.frx":0000
   ScaleHeight     =   8460
   ScaleWidth      =   13875
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   6075
      Left            =   10920
      Picture         =   "puz_help.frx":25DFB
      ScaleHeight     =   6015
      ScaleWidth      =   4035
      TabIndex        =   10
      Top             =   3360
      Width           =   4095
   End
   Begin VB.PictureBox Picture3 
      Height          =   2295
      Left            =   11640
      Picture         =   "puz_help.frx":2CC9A
      ScaleHeight     =   2235
      ScaleWidth      =   2235
      TabIndex        =   9
      Top             =   960
      Width           =   2295
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   840
      Picture         =   "puz_help.frx":32A47
      ScaleHeight     =   1515
      ScaleWidth      =   1515
      TabIndex        =   8
      Top             =   6600
      Width           =   1575
   End
   Begin VB.ListBox lsthelp 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   6060
      ItemData        =   "puz_help.frx":3A949
      Left            =   240
      List            =   "puz_help.frx":3A956
      TabIndex        =   4
      Top             =   3360
      Width           =   3975
   End
   Begin VB.CommandButton cmdindex 
      BackColor       =   &H00FFFFFF&
      Caption         =   "INDEX"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdsrch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtsrch 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton cmdgo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome !  It's A Puzzle Game Help"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   360
      Width           =   7935
   End
   Begin VB.Label lblhelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   6015
      Left            =   4560
      TabIndex        =   6
      Top             =   3360
      Width           =   6045
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Type In The Text To Find"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   3615
   End
End
Attribute VB_Name = "frmpuzhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdgo_Click()
Dim h As Integer
h = 0
lblhelp.Caption = ""
If txtsrch.Text <> "" Then
  If InStr(UCase(puzobject), " " + UCase(txtsrch.Text) + " ") Then
    lblhelp.Caption = puzobject
    h = h + 1
  End If
  If InStr(UCase(puzplaying), " " + UCase(txtsrch.Text) + " ") Then
    lblhelp.Caption = lblhelp.Caption + puzplaying
    h = h + 1
  End If
  If InStr(UCase(puzscoring), " " + UCase(txtsrch.Text) + " ") Then
    lblhelp.Caption = lblhelp.Caption + puzscoring
    h = h + 1
  End If
End If
If h = 0 Then
   lblhelp.Caption = "Couldn't Find The Matched Text"
End If
End Sub

Private Sub cmdindex_Click()
txtsrch.Enabled = False
lsthelp.Enabled = True
cmdgo.Enabled = False
End Sub

Private Sub cmdsrch_Click()
lsthelp.Enabled = False
txtsrch.Enabled = True
cmdgo.Enabled = True
End Sub

Private Sub Form_Load()

End Sub

Private Sub lsthelp_Click()
If lsthelp.Selected(0) Then
    lblhelp.Caption = puzobject
ElseIf lsthelp.Selected(2) Then
    lblhelp.Caption = puzplaying
Else
   lblhelp.Caption = ""
End If
End Sub

