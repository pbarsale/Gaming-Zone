VERSION 5.00
Begin VB.Form frmpuz2 
   Caption         =   "Level2  Puzzle"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14280
   LinkTopic       =   "Form7"
   Picture         =   "puzz_sub.frx":0000
   ScaleHeight     =   8280
   ScaleWidth      =   14280
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdback5 
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
      Picture         =   "puzz_sub.frx":23BE4
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   24
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   23
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   22
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   21
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   20
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   19
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   18
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   17
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   12
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   11
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmddigits1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Caption         =   "Puzzle"
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
      Height          =   5295
      Left            =   4320
      TabIndex        =   24
      Top             =   2040
      Width           =   7335
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   480
         Picture         =   "puzz_sub.frx":24B05
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3720
         Width           =   1575
      End
      Begin VB.CommandButton cmdnewgame 
         BackColor       =   &H00FFFFFF&
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
         Height          =   975
         Left            =   480
         Picture         =   "puzz_sub.frx":25347
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdhelp 
         BackColor       =   &H00FFFFFF&
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
         Height          =   975
         Left            =   480
         Picture         =   "puzz_sub.frx":25C41
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   750
         Left            =   2880
         Top             =   480
      End
      Begin VB.Label lbltime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3840
         TabIndex        =   25
         Top             =   480
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
Attribute VB_Name = "frmpuz2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim time As Integer
Dim starttime


Private Sub cmdback5_Click()
Me.Hide
frmsub2.Show
End Sub

Private Sub cmddigits1_Click(Index As Integer)

If time = 0 Then
  starttime = Now
  Timer1.Enabled = True
  time = 1
End If
  
Dim i, j As Integer
Dim c As Integer
c = 0
i = 4
j = 10

If Index < 5 Then

     If cmddigits1(Index + 5).Caption = "" Then
         c = xchange(Index + 5, Index)
         GoTo label
     ElseIf cmddigits1(Index + 1).Caption = "" And (Index + 1) < 5 Then
         c = xchange(Index + 1, Index)
         GoTo label
     ElseIf (Index - 1) > -1 Then
         If cmddigits1(Index - 1).Caption = "" Then
          c = xchange(Index - 1, Index)
          GoTo label
         End If
     End If
         
         
ElseIf Index > 19 Then
 
     If cmddigits1(Index - 5).Caption = "" Then
         c = xchange(Index - 5, Index)
         GoTo label
     ElseIf cmddigits1(Index - 1).Caption = "" And (Index - 1) > 19 Then
         c = xchange(Index - 1, Index)
         GoTo label
     ElseIf (Index + 1) <= 24 Then
       If cmddigits1(Index + 1).Caption = "" Then
          c = xchange(Index + 1, Index)
          GoTo label
       End If
     End If
         
Else
   
   Do While i < 15 And j < 21
   
   If Index > i And Index < j Then
   
     If cmddigits1(Index + 5).Caption = "" Then
        c = xchange(Index + 5, Index)
        GoTo label
     ElseIf cmddigits1(Index - 5).Caption = "" Then
        c = xchange(Index - 5, Index)
        GoTo label
     ElseIf cmddigits1(Index + 1).Caption = "" And (Index + 1) < j Then
        c = xchange(Index + 1, Index)
        GoTo label
     ElseIf cmddigits1(Index - 1).Caption = "" And (Index - 1) > i Then
        c = xchange(Index - 1, Index)
        GoTo label
     End If
   End If

   i = i + 5
   j = j + 5
   
   Loop
  End If
   
label:
   tally
 
        
      
End Sub
 

Private Sub cmdexit_Click()
mnuexit_Click
End Sub

Private Sub cmdhelp_Click()
mnucontents_Click
End Sub

Private Sub cmdnewgame_Click()
mnunew_Click
End Sub

Private Sub Form_Load()
lbltime.Caption = Format$("00  00  00")
time = 0
setcmd
End Sub
Public Sub setcmd()

cmddigits1(0).Caption = "4"
cmddigits1(1).Caption = "16"
cmddigits1(2).Caption = "11"
cmddigits1(3).Caption = "20"
cmddigits1(4).Caption = "3"
cmddigits1(5).Caption = "12"
cmddigits1(6).Caption = "21"
cmddigits1(7).Caption = ""
cmddigits1(8).Caption = "8"
cmddigits1(9).Caption = "13"
cmddigits1(10).Caption = "7"
cmddigits1(11).Caption = "1"
cmddigits1(12).Caption = "10"
cmddigits1(13).Caption = "23"
cmddigits1(14).Caption = "5"
cmddigits1(15).Caption = "22"
cmddigits1(16).Caption = "18"
cmddigits1(17).Caption = "14"
cmddigits1(18).Caption = "2"
cmddigits1(19).Caption = "17"
cmddigits1(20).Caption = "15"
cmddigits1(21).Caption = "6"
cmddigits1(22).Caption = "24"
cmddigits1(23).Caption = "19"
cmddigits1(24).Caption = "9"

End Sub
Public Function xchange(index1 As Integer, index2 As Integer) As Integer
Dim temp As String

temp = cmddigits1(index1).Caption
cmddigits1(index1).Caption = cmddigits1(index2).Caption
cmddigits1(index2).Caption = temp
xchange = 1
End Function

Public Sub tally()
Dim k As Integer

Dim count As Integer
count = 0
For k = 0 To 24
  
  If cmddigits1(k).Caption = CStr(k + 1) Then
     count = count + 1
  End If
  
Next k

If count = 24 Then
   successful
End If
End Sub

Public Sub successful()
Dim result As String
cmddigits1(24).Caption = "25"
MsgBox (" Congratulations ! You Win ")


Timer1.Enabled = False
result = MsgBox("Do You Want To Continue", vbYesNo)
If result = vbYes Then
   mnunew_Click
Else
  Me.Hide
  frmlist.Show
End If



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
time = 0
lbltime.Caption = Format$("00   00   00")
setcmd
End Sub
