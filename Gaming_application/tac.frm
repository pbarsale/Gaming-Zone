VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmtictac 
   BorderStyle     =   0  'None
   Caption         =   "Tic  Tac  Toe"
   ClientHeight    =   9075
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   14910
   FillColor       =   &H000000C0&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "tac.frx":0000
   ScaleHeight     =   16304.06
   ScaleMode       =   0  'User
   ScaleWidth      =   41207.02
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   14280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "tac.frx":23BDB
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8640
      Width           =   1575
   End
   Begin VB.CheckBox chkp2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   14280
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CheckBox chkp1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   14280
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdstart 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "tac.frx":2435C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Picture         =   "tac.frx":24ADD
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtpoints 
      DataField       =   "points"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   6600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   480
      Top             =   7440
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"tac.frx":259FE
      OLEDBString     =   $"tac.frx":25A90
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "gaming"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmddig 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Player's Turn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   13800
      TabIndex        =   6
      Top             =   360
      Width           =   2895
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
      Begin VB.Menu mnucontent 
         Caption         =   "Contents"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmtictac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim player1 As String
Dim player2 As String
Dim pass1 As String
Dim pass2 As String
Public loaded As Boolean
Dim winner As String
Dim i As Integer
Dim c As Integer


Private Sub cmddig_Click(Index As Integer)
Dim count As Integer
count = 0
If cmddig(Index).Caption = "" Then
  If c = 0 Then
    cmddig(Index).Caption = "X"
    c = 1
    chkp1.Value = 0
    chkp2.Value = 1
  ElseIf c = 1 Then
    cmddig(Index).Caption = "0"
    c = 0
    chkp1.Value = 1
    chkp2.Value = 0
  End If
End If

If cmddig(0).Caption = cmddig(1).Caption And cmddig(0).Caption = cmddig(2).Caption And (cmddig(0).Caption = "X" Or cmddig(0).Caption = "0") Then
   terminate
ElseIf cmddig(0).Caption = cmddig(3).Caption And cmddig(0).Caption = cmddig(6).Caption And (cmddig(0).Caption = "X" Or cmddig(0).Caption = "0") Then
   terminate
ElseIf cmddig(0).Caption = cmddig(4).Caption And cmddig(0).Caption = cmddig(8).Caption And (cmddig(0).Caption = "X" Or cmddig(0).Caption = "0") Then
   terminate
ElseIf cmddig(1).Caption = cmddig(4).Caption And cmddig(1).Caption = cmddig(7).Caption And (cmddig(1).Caption = "X" Or cmddig(1).Caption = "0") Then
   terminate
ElseIf cmddig(2).Caption = cmddig(4).Caption And cmddig(2).Caption = cmddig(6).Caption And (cmddig(2).Caption = "X" Or cmddig(2).Caption = "0") Then
   terminate
ElseIf cmddig(2).Caption = cmddig(5).Caption And cmddig(2).Caption = cmddig(8).Caption And (cmddig(2).Caption = "X" Or cmddig(2).Caption = "0") Then
   terminate
ElseIf cmddig(3).Caption = cmddig(4).Caption And cmddig(3).Caption = cmddig(5).Caption And (cmddig(3).Caption = "X" Or cmddig(3).Caption = "0") Then
   terminate
ElseIf cmddig(6).Caption = cmddig(7).Caption And cmddig(6).Caption = cmddig(8).Caption And (cmddig(6).Caption = "X" Or cmddig(6).Caption = "0") Then
   terminate
Else
  For i = 0 To 8
    If cmddig(i).Caption <> "" Then
      count = count + 1
    End If
   If count = 9 Then
     draw
   End If
  Next i
End If
End Sub

Private Sub cmdstart_Click()
For i = 0 To 8 Step 1
cmddig(i).Enabled = True
Next i


 MsgBox (" All The Very Best To Both Of You ")
chkp1.Value = 1
chkp2.Value = 0
cmdstart.Enabled = False
End Sub

Private Sub Command1_Click()
Me.Hide
frmsub1.Show
End Sub

Private Sub Command2_Click()
Me.Hide
frmlist.Show
End Sub

Private Sub Form_Load()

chkp1.Value = 0
chkp2.Value = 0
frmlist.Hide
frmtictac.Show
c = 0
start
For i = 1 To 8 Step 1
  Load cmddig(i)
  If (i = 3) Then
    cmddig(i).Top = cmddig(0).Top + 1800
  ElseIf (i = 6) Then
    cmddig(i).Top = cmddig(3).Top + 1800
  Else
    cmddig(i).Top = cmddig(i - 1).Top
    cmddig(i).Left = cmddig(i - 1).Left + 2800
  End If
  cmddig(i).Visible = True
Next i
 cmddig(0).Visible = True
loaded = True
End Sub

Private Sub mnucontent_Click()

frmtichelp.Show
frmtichelp.lblhelp.Caption = ""
frmtichelp.lsthelp.Enabled = False
frmtichelp.txtsrch.Enabled = False
frmtichelp.cmdgo.Enabled = False
End Sub

Public Sub terminate()

 If chkp2.Value Then
     winner = player1
 ElseIf chkp1.Value Then
    winner = player2
 End If

  win
End Sub

Public Sub draw()
Dim result As String
result = MsgBox(" There Is A Tie ,Do You Want To Play Again ", vbYesNo)
If result = vbYes Then
  mnunew_Click
Else
  Me.Hide
  frmlist.Show
End If
End Sub
 
Private Sub mnuexit_Click()
Me.Hide
frmlist.Show
End Sub

 Public Sub mnunew_Click()
c = 0
For i = 0 To 8
   chkp1.Value = 0
   chkp2.Value = 0
   cmddig(i).Caption = ""
   cmddig(i).Enabled = False
Next i
start
End Sub

Private Sub mnusp2_Click()
End
End Sub

Private Sub start()

lbl1:

player1 = InputBox("First Player - Welcome To The Tic Tac Toe Game                                                                                                                         Enter Your UserName ", "Tic Tac Toe - First Player")
pass1 = InputBox("First Player - Enter Your Password ")

If player1 = "" Or pass1 = "" Then GoTo lbl1

 If Not (confirm(player1, pass1)) Then
 MsgBox (" Invalid Username Or Password ,Please Try Again")
 GoTo lbl1
End If
  
  chkp1.Caption = player1
 
lbl2:
player2 = InputBox("Second Player - Welcome To The Tic Tac Toe Game                                                                                                                         Enter Your UserName ", "Tic Tac Toe - Second Player")
pass2 = InputBox("Second Player - Enter Your Password")

If player2 = "" Or pass2 = "" Then GoTo lbl2

If Not (confirm(player2, pass2)) Then
 MsgBox (" Invalid Username Or Password ,Please Try Again")
 GoTo lbl2
End If

 chkp2.Caption = player2
 chkp1.Value = 1
 
MsgBox (player1 & " The Icon Alloted To You Is 'X'  ." & player2 & " The Icon Alloted To You Is '0'")
cmdstart.Enabled = True
End Sub
Private Function confirm(s1 As String, s2 As String) As Boolean

Dim dbookmark As Double
Dim db1 As Double
Dim db2 As Double

dbookmark = Adodc1.Recordset.Bookmark
Dim query1 As String
Dim query2 As String

query1 = "username like '" & s1 & "'"
query2 = "password like '" & s2 & "'"
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find query1

If Adodc1.Recordset.EOF Then
   confirm = 0
Else
  db1 = Adodc1.Recordset.Bookmark
  Adodc1.Recordset.Find query2
  If Adodc1.Recordset.EOF Then
    confirm = 0
  Else
    db2 = Adodc1.Recordset.Bookmark
    If db1 - db2 = 0 Then
     confirm = 1
    Else
      confirm = 0
    End If
  End If
End If
 Adodc1.Recordset.Bookmark = dbookmark
End Function

Public Sub win()
Dim result As String

Dim dbookmark As Double
dbookmark = Adodc1.Recordset.Bookmark
Dim query As String

query = "username like '" & winner & "'"
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find query

If Adodc1.Recordset.EOF Then
   Adodc1.Recordset.Bookmark = dbookmark
End If
   
txtpoints.Text = txtpoints.Text + 20
Adodc1.Recordset.Update ("points"), (txtpoints.Text)

result = MsgBox("Congratulation " + winner + " !!!!!  You Have Won ,Your Total Ponits Are " + txtpoints.Text + " ,Do You Want To Play Again", vbYesNo)
 If result = vbYes Then
   mnunew_Click
  Else
   Me.Hide
   frmlist.Show
 End If
End Sub

