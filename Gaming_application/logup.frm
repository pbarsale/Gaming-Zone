VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmlogup 
   Caption         =   "Log Up Form"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   14910
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   Picture         =   "logup.frx":0000
   ScaleHeight     =   8760
   ScaleWidth      =   14910
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdback2 
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
      Picture         =   "logup.frx":59B50
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtgender 
      DataField       =   "gender"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   10440
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.OptionButton optfemale 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8520
      TabIndex        =   13
      Top             =   2640
      Width           =   1695
   End
   Begin VB.OptionButton optmale 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6360
      TabIndex        =   12
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CREATE  MY ACCOUNT"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8880
      Width           =   3735
   End
   Begin VB.TextBox txtaddress 
      Alignment       =   2  'Center
      DataField       =   "address"
      DataSource      =   "Adodc1"
      Height          =   1545
      Left            =   6240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3840
      Width           =   4095
   End
   Begin VB.TextBox txtuname 
      Alignment       =   2  'Center
      DataField       =   "username"
      DataSource      =   "Adodc1"
      Height          =   465
      Left            =   6240
      TabIndex        =   3
      Top             =   5760
      Width           =   4095
   End
   Begin VB.TextBox txtpassword 
      Alignment       =   2  'Center
      DataField       =   "password"
      DataSource      =   "Adodc1"
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   6240
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   6480
      Width           =   4095
   End
   Begin VB.TextBox txtphone 
      Alignment       =   2  'Center
      DataField       =   "phone"
      DataSource      =   "Adodc1"
      Height          =   465
      Left            =   6240
      TabIndex        =   1
      Top             =   7200
      Width           =   4095
   End
   Begin VB.TextBox txtname 
      Alignment       =   2  'Center
      DataField       =   "pname"
      DataSource      =   "Adodc1"
      Height          =   465
      Left            =   6240
      TabIndex        =   0
      Top             =   1560
      Width           =   4095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   3240
      Top             =   8760
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1508
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
      Connect         =   $"logup.frx":5AA71
      OLEDBString     =   $"logup.frx":5AB03
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "gaming"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GENDER"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   2400
      Width           =   4215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   3960
      Width           =   4215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   5760
      Width           =   4215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   6480
      Width           =   4215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PHONE NO."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   7200
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1560
      Width           =   4215
   End
End
Attribute VB_Name = "frmlogup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gblnaddmode As Boolean
Public gblbookmark


Private Sub cmdback2_Click()
Me.Hide
frmintro.Show
End Sub

Private Sub cmdOK_Click()
On Error GoTo errorhandler:

  If optmale.Value Then
    txtgender.Text = "M"
  ElseIf optfemale.Value Then
    txtgender.Text = "F"
  End If

  If gblnaddmode Then
    If Adodc1.Recordset.BOF Or Adodc1.Recordset.EOF Then
       Exit Sub
    End If
  Adodc1.Recordset.Update
  gblnaddmode = False
  
  MsgBox (" Your Account Has Been Successfully Created ")
  frmlogup.Hide
  frmlist.Show
  Exit Sub
 End If


errorhandler:
 MsgBox ("Username Already Exist Or Wrong Information Entered In Either Fields ,Please Try Again ")
 
 End Sub

'Function search() As Integer


'Dim fname As String
'Dim sfindcriterion As String
'Dim result As Integer
'sfindcriterion = "username like '" & fname & "'"
'Adodc1.Recordset.MoveFirst
'Adodc1.Recordset.Find sfindcriterion

'If Adodc1.Recordset.EOF Then
  ' result = 0
'Else
   'result = 1
'End If

'search = result

'End Function
Private Sub Label8_Click()

End Sub

