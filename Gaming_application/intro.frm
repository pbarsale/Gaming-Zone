VERSION 5.00
Begin VB.Form frmintro 
   BackColor       =   &H00FFFF00&
   Caption         =   "Home Page"
   ClientHeight    =   8760
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   15465
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   Picture         =   "intro.frx":0000
   ScaleHeight     =   8760
   ScaleWidth      =   15465
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Height          =   975
      Left            =   120
      Picture         =   "intro.frx":75EA1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdlogup 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log Up"
      Height          =   1335
      Left            =   3600
      Picture         =   "intro.frx":76DC2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdlogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOG IN"
      Height          =   1335
      Index           =   0
      Left            =   600
      Picture         =   "intro.frx":776BC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "EXIT"
      Height          =   1335
      Left            =   6840
      Picture         =   "intro.frx":781ED
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "To Create A New Account"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   5880
      Width           =   2775
   End
   Begin VB.OLE OLE1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   0  'Transparent
      Class           =   "Word.Document.12"
      Height          =   1140
      Left            =   5640
      OleObjectBlob   =   "intro.frx":787DF
      SourceDoc       =   "C:\Documents and Settings\Administrator\Desktop\Doc2.docx"
      TabIndex        =   3
      Top             =   1200
      Width           =   8055
   End
End
Attribute VB_Name = "frmintro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c As Integer
Private Sub cmdexit_Click()
Me.Hide
frmstart.Show
End Sub

Private Sub cmdlogin_Click(Index As Integer)
frmlist.Show
Me.Hide
End Sub

Private Sub cmdlogup_Click()

frmlogup.Show
frmlogup.optfemale.Value = False
frmlogup.optmale.Value = False
frmlogup.gblnaddmode = True

If c = 0 Then
frmlogup.Adodc1.Recordset.MoveLast
frmlogup.Adodc1.Recordset.AddNew

c = 1
End If
Me.Hide
End Sub

Private Sub cmdnext_Click()

End Sub

Private Sub Command1_Click()
Me.Hide
frmstart.Show
End Sub

Private Sub Form_Load()
c = 0
frmsupport.Show
End Sub
