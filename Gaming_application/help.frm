VERSION 5.00
Begin VB.Form frmtichelp 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tic Tac Toe Help"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   14355
   LinkTopic       =   "Form2"
   Picture         =   "help.frx":0000
   ScaleHeight     =   8895
   ScaleWidth      =   14355
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   11400
      Picture         =   "help.frx":41759
      ScaleHeight     =   1275
      ScaleWidth      =   1995
      TabIndex        =   9
      Top             =   1800
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   10680
      Picture         =   "help.frx":42616
      ScaleHeight     =   4755
      ScaleWidth      =   4875
      TabIndex        =   8
      Top             =   3360
      Width           =   4935
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
      TabIndex        =   7
      Top             =   2640
      Width           =   495
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
      TabIndex        =   5
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton cmdsrch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton cmdindex 
      BackColor       =   &H00FFFFFF&
      Caption         =   "INDEX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
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
      ItemData        =   "help.frx":4F9AC
      Left            =   240
      List            =   "help.frx":4F9BF
      TabIndex        =   0
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Width           =   3375
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
      Height          =   6615
      Left            =   4440
      TabIndex        =   2
      Top             =   3360
      Width           =   6045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome !  It's A Tic  Tac  Toe  Game Help"
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
      Height          =   975
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   9495
   End
End
Attribute VB_Name = "frmtichelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim h As Integer

Private Sub cmdgo_Click()
h = 0
lblhelp.Caption = ""
If txtsrch.Text <> "" Then
  If InStr(UCase(ticobject), " " + UCase(txtsrch.Text) + " ") Then
    lblhelp.Caption = ticobject
    h = h + 1
  End If
  If InStr(UCase(ticplaying), " " + UCase(txtsrch.Text) + " ") Then
    lblhelp.Caption = lblhelp.Caption + ticplaying
    h = h + 1
  End If
  If InStr(UCase(ticscoring), " " + UCase(txtsrch.Text) + " ") Then
    lblhelp.Caption = lblhelp.Caption + ticscoring
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


Private Sub lsthelp_Click()

If lsthelp.Selected(0) Then
    lblhelp.Caption = ticobject
ElseIf lsthelp.Selected(2) Then
    lblhelp.Caption = ticplaying
ElseIf lsthelp.Selected(4) Then
    lblhelp.Caption = ticscoring
Else
   lblhelp.Caption = ""
End If

End Sub

