VERSION 5.00
Begin VB.Form frmpuz1 
   Caption         =   "Level1 Puzzle"
   ClientHeight    =   8880
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13980
   LinkTopic       =   "Form6"
   Picture         =   "puzzle_main.frx":0000
   ScaleHeight     =   8880
   ScaleWidth      =   13980
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdback4 
      BackColor       =   &H00FFFFFF&
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
      Picture         =   "puzzle_main.frx":2D510
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdhelp 
      BackColor       =   &H0000FFFF&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton cmdnewgame 
      BackColor       =   &H0000FFFF&
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H0000FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   12
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   11
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmddigits 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   6960
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      Caption         =   "PUZZLE 1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5535
      Left            =   4320
      TabIndex        =   20
      Top             =   2160
      Width           =   6495
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   2760
         Top             =   1200
      End
      Begin VB.Label lbltime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3600
         TabIndex        =   21
         Top             =   1200
         Width           =   2295
      End
   End
   Begin VB.Menu mnugame 
      Caption         =   "Game"
      Begin VB.Menu mnunew 
         Caption         =   "New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnusp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnucontents 
         Caption         =   "Contents"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmpuz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public time As Integer
Dim idx As Integer
Dim starttime

Private Sub cmdback4_Click()
Me.Hide
frmsub2.Show
End Sub

Private Sub cmddigits_Click(Index As Integer)
If time = 0 Then
  starttime = Now
  Timer1.Enabled = True
  time = 1
End If
  
Dim i, j As Integer
Dim c As Integer
c = 0
i = 3
j = 8
idx = Index

If idx < 4 Then

     If cmddigits(idx + 4).Caption = "" Then
         c = swap(idx + 4, idx)
         GoTo label
     ElseIf cmddigits(idx + 1).Caption = "" And (idx + 1) < 4 Then
         c = swap(idx + 1, idx)
         GoTo label
     ElseIf (idx - 1) > -1 Then
         If cmddigits(idx - 1).Caption = "" Then
          c = swap(idx - 1, idx)
          GoTo label
         End If
     End If
         
         
ElseIf idx > 11 Then
 
     If cmddigits(idx - 4).Caption = "" Then
         c = swap(idx - 4, idx)
         GoTo label
     ElseIf cmddigits(idx - 1).Caption = "" And (idx - 1) > 11 Then
         c = swap(idx - 1, idx)
         GoTo label
     ElseIf (idx + 1) <= 15 Then
       If cmddigits(idx + 1).Caption = "" Then
          c = swap(idx + 1, idx)
          GoTo label
       End If
     End If
         
Else
   
   Do While i < 8 And j < 13
   
   If idx > i And idx < j Then
   
     If cmddigits(idx + 4).Caption = "" Then
        c = swap(idx + 4, idx)
        GoTo label
     ElseIf cmddigits(idx - 4).Caption = "" Then
        c = swap(idx - 4, idx)
        GoTo label
     ElseIf cmddigits(idx + 1).Caption = "" And (idx + 1) < j Then
        c = swap(idx + 1, idx)
        GoTo label
     ElseIf cmddigits(idx - 1).Caption = "" And (idx - 1) > i Then
        c = swap(idx - 1, idx)
        GoTo label
     End If
   End If

   i = i + 4
   j = j + 4
   
   Loop
  End If
   
label:
   comp
 
        
      
End Sub

Public Function swap(idx1 As Integer, idx2 As Integer) As Integer
Dim temp As String

temp = cmddigits(idx1).Caption
cmddigits(idx1).Caption = cmddigits(idx2).Caption
cmddigits(idx2).Caption = temp
swap = 1

End Function

Private Sub comp()
Dim k As Integer

Dim count As Integer
count = 0
For k = 0 To 14
  
  If cmddigits(k).Caption = CStr(k + 1) Then
     count = count + 1
  End If
  
Next k

If count = 15 Then
   win
End If

End Sub

Private Sub win()
Dim result As String
cmddigits(15).Caption = "16"


Timer1.Enabled = False
result = MsgBox(" Congratulations ! You Win Do You Want To Continue", vbYesNo)
If result = vbYes Then
   mnunew_Click
Else
  Me.Hide
  frmlist.Show
End If
End Sub







Private Sub cmdexit_Click()

   Me.Hide
   frmlist.Show
End Sub

Private Sub cmdhelp_Click()
mnucontents_Click
End Sub

Private Sub cmdnewgame_Click()
mnunew_Click
End Sub

Private Sub Form_Load()
begin
End Sub


Private Sub setcmd()
cmddigits(0).Caption = "6"
cmddigits(1).Caption = "1"
cmddigits(2).Caption = "9"
cmddigits(3).Caption = "7"
cmddigits(4).Caption = "14"
cmddigits(5).Caption = "12"
cmddigits(6).Caption = "2"
cmddigits(7).Caption = "13"
cmddigits(8).Caption = "5"
cmddigits(9).Caption = ""
cmddigits(10).Caption = "11"
cmddigits(11).Caption = "3"
cmddigits(12).Caption = "10"
cmddigits(13).Caption = "15"
cmddigits(14).Caption = "4"
cmddigits(15).Caption = "8"
End Sub

Private Sub mnucontents_Click()
frmpuzhelp.Show
End Sub

Private Sub mnuexit_Click()
Me.Hide
frmlist.Show
End Sub

Private Sub mnunew_Click()
begin

End Sub

Private Sub Timer1_Timer()
lbltime.Caption = Format$(Now - starttime, " hh : mm : ss ")
End Sub

Public Sub begin()
Timer1.Enabled = False
lbltime.Caption = Format$("00  00  00")
time = 0
setcmd
End Sub
