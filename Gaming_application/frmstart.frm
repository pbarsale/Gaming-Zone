VERSION 5.00
Begin VB.Form frmstart 
   Caption         =   "Introduction "
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   Picture         =   "frmstart.frx":0000
   ScaleHeight     =   8385
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00FFFF80&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7560
      Width           =   2295
   End
   Begin VB.CommandButton cmdproceed 
      BackColor       =   &H00FFFF80&
      Caption         =   "Proceed"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   2415
   End
   Begin VB.OptionButton optowner 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Owner"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   3120
      Width           =   3615
   End
   Begin VB.OptionButton optcustomer 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   3120
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Your Account Type :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   4815
   End
End
Attribute VB_Name = "frmstart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdproceed_Click()
  frmsupport.Show
If optowner.Value Then
  Me.Hide
  frmLogin.Show
  frmLogin.txtpassword = ""
  frmLogin.txtUserName = ""
  
ElseIf optcustomer.Value Then
  Me.Hide
  frmintro.Show
Else
  MsgBox ("Please Select One Option ")
End If

End Sub


