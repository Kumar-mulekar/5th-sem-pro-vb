VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   0  'None
   Caption         =   "User"
   ClientHeight    =   7740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12450
   LinkTopic       =   "Form4"
   ScaleHeight     =   7740
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      Height          =   735
      Left            =   1560
      TabIndex        =   12
      Top             =   6600
      Width           =   9135
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   1560
      TabIndex        =   6
      Top             =   3720
      Width           =   9135
      Begin VB.TextBox Text9 
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Text            =   "Which country do you want to visit ? "
         Top             =   1680
         Width           =   8415
      End
      Begin VB.TextBox Text8 
         Height          =   615
         Left            =   6360
         TabIndex        =   10
         Text            =   "Re-enter Password"
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         Height          =   615
         Left            =   3360
         TabIndex        =   9
         Text            =   "Password"
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Text            =   "User Name"
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Text            =   "Select Access"
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   7920
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "Phone No"
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Text            =   "Email"
      Top             =   2520
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Text            =   "Address"
      Top             =   1680
      Width           =   8415
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   6600
      TabIndex        =   2
      Text            =   "Last Name"
      Top             =   840
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   9135
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Text            =   "First Name"
         Top             =   360
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim id As Integer

Private Sub Command1_Click()
If Text7.Text <> Text8.Text Then
    MsgBox "Password does not match", vbInformation
    Exit Sub
End If
Dim us As ADODB.Recordset
Set us = New ADODB.Recordset
us.Open "select *from login", con, adOpenDynamic, adLockPessimistic, adCmdText

Call Auto_num
With us
    .AddNew
    .Fields(0).Value = id
    .Fields(1).Value = Text6.Text
    .Fields(2).Value = Text7.Text
    .Fields(3).Value = Text1.Text
    .Fields(4).Value = Text2.Text
    .Fields(5).Value = Text3.Text
    .Fields(6).Value = Text4.Text
    .Fields(7).Value = CLng(Text5.Text)
    .Fields(8).Value = Text9.Text
    .Fields(9).Value = Combo1.Text
    .Update
End With
MsgBox "New user Added", vbMsgBoxRight
us.Close
Call clearText
End Sub
Private Sub clearText()
Text1.Text = "First Name"
Text2.Text = "Last Name"
Text3.Text = "Address"
Text4.Text = "Email"
Text5.Text = "Phone No"
Combo1.Text = "Select Access"
Text6.Text = "User Name"
Text7.Text = "Password"
Text8.Text = "Re-enter Password"
Text9.Text = "Which country do you want to visit ? "
End Sub
Private Sub Auto_num()
   Dim auto As ADODB.Recordset
   Set auto = New ADODB.Recordset
   auto.Open "select *from login", con, adOpenDynamic, adLockPessimistic, adCmdText
    With auto
        If .RecordCount = 0 Then
           id = 1
        Else
           .MoveLast
            id = (auto!id) + 1
        End If
        
        auto.Close
    End With
End Sub

Private Sub Form_Load()
'***form location
With frmUser
.BackColor = RGB(238, 238, 238)
.Top = frmmain.Top + 1000
.Left = frmmain.Left + 3735
.Height = frmmain.Height - 1000
.Width = frmmain.Width - 3735
End With


'""""""""""combo1
Combo1.AddItem "Admin"
Combo1.AddItem "User"

End Sub

'++++++++++++++++++++++++++++++++++++++textbox validation
Private Sub Text1_KeyPress(KeyAscii As Integer)
Call validateA(KeyAscii)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
Call validateA(KeyAscii)
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
Call validateE(KeyAscii)
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
Call validateN(KeyAscii)
End Sub

'+++++++++++++++++++++++++++++++++++++++++end textbox validation
'===================================textbox clear
Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text1.Text = ""
End Sub

Private Sub Text2_GotFocus()
Text2.Text = ""
End Sub
Private Sub Text3_GotFocus()
Text3.Text = ""
End Sub
Private Sub Text4_GotFocus()
Text4.Text = ""
End Sub
Private Sub Text5_GotFocus()
Text5.Text = ""
End Sub
Private Sub Text6_GotFocus()
Text6.Text = ""
End Sub
Private Sub Text7_GotFocus()
Text7.Text = ""
End Sub
Private Sub Text8_GotFocus()
Text8.Text = ""
End Sub

Private Sub Text9_GotFocus()
Text9.Text = ""
End Sub
'=============================================== end textbox clear
