VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "login"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10500
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmLogin.frx":54AA
   ScaleHeight     =   5880
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":12F79
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":13010
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":13116
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   4800
   End
   Begin VB.Frame Frame3 
      Height          =   4815
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   4215
      Begin VB.CommandButton btnlogin 
         Caption         =   "Log In"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         TabIndex        =   9
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2160
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   2160
         TabIndex        =   6
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4815
      Left            =   5520
      TabIndex        =   1
      Top             =   360
      Width           =   4215
      Begin VB.CommandButton btnresetpass 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   20
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   19
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   2400
         TabIndex        =   17
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   2400
         TabIndex        =   16
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   2400
         TabIndex        =   15
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Enter"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   9735
      Begin VB.Label loginlbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Log In"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         TabIndex        =   3
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label resetpass 
         BackStyle       =   0  'Transparent
         Caption         =   "Reset Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5880
         TabIndex        =   2
         Top             =   1440
         Width           =   3495
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public tm, flef, whichframe As Integer
Public pRecordset As ADODB.Recordset
Private Sub btnlogin_Click()
Dim flag, log As Boolean
flag = True


pRecordset.MoveFirst
While Not pRecordset.EOF = True
  If pRecordset.Fields(1).Value = Text1.Text And pRecordset.Fields(2).Value = Text2.Text Then
     Form1.Hide
     userName = pRecordset.Fields(1).Value
     userAccess = pRecordset.Fields(9).Value
     frmmain.Show
     flag = False
     Unload Me
  End If
  pRecordset.MoveNext
  
Wend
If flag = True Then
   MsgBox ("Invalid Details")
End If

End Sub

Private Sub btnresetpass_Click()
Dim flag As Boolean

flag = True
pRecordset.MoveFirst
While Not pRecordset.EOF = True
  If pRecordset.Fields("fname").Value = Text3.Text And pRecordset.Fields("lname").Value = Text4.Text And pRecordset.Fields("quetion").Value = Text5.Text Then
     flag = False
     If (Text6.Text = Text7.Text) Then
        pRecordset.Fields("password").Value = Text6.Text
        pRecordset.Update
        MsgBox ("DONE!!!!")
     Else
        MsgBox ("Password Does Not Match")
     End If
  End If
  pRecordset.MoveNext
Wend
If flag = True Then
   MsgBox ("Invalid Details")
End If
End Sub

Private Sub Form_Load()

Form1.Picture = Nothing
'skn.LoadSkin App.Path + "\Styles\Office2010.cjstyles", ""
'Office2007 Office2010  WinXP.Luna WinXP.Royale Codejock
'skn.ApplyWindow Me.hWnd
Frame2.Visible = False


'database
Call Module2.main
Set pRecordset = New ADODB.Recordset
pRecordset.Open "select *from login", con, adOpenDynamic, adLockPessimistic, adCmdText
End Sub


Private Sub resetpass_Click()
Frame2.Left = 960
Frame2.Visible = True
Frame3.Visible = False

tm = 960
flef = 5520
Timer1.Enabled = True
whichframe = 1
End Sub

Private Sub loginlbl_Click()
Frame3.Left = 5520
Frame2.Visible = False
Frame3.Visible = True

tm = 5520
flef = 960

Timer1.Enabled = True
whichframe = 2
End Sub

Private Sub Timer1_Timer()
If whichframe = 1 Then
 If (tm <= flef) Then
    With Frame2
      .Move (tm)
    End With
    tm = tm + 30
 End If
 If tm >= flef Then
  Timer1.Enabled = False
 End If
ElseIf whichframe = 2 Then
  If (tm >= flef) Then
    With Frame3
      .Move (tm)
    End With
    tm = tm - 30
 End If
 If tm <= flef Then
  Timer1.Enabled = False
 End If
End If
End Sub
