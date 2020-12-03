VERSION 5.00
Begin VB.Form frmSupplier 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12570
   LinkTopic       =   "Form2"
   ScaleHeight     =   8205
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   9
      Top             =   6000
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   6000
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2880
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   7680
      TabIndex        =   2
      Text            =   "Last Name"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Text            =   "First Name"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   7800
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "Ph.no"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Text            =   "E-mail"
      Top             =   4200
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Text            =   "Address"
      Top             =   3360
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Text            =   "S-ID"
      Top             =   2520
      Width           =   1935
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recS As ADODB.Recordset
Private Sub Combo1_click()
    If Combo1.Text = "EDIT" Or Combo1.Text = "DELETE" Then
        MsgBox "Enter First and Last Name then hit Enter."
        Command4.Visible = True
    Else
        Command4.Visible = False
    End If
    
    
    'enable and disable button
     If Combo1.Text = "ADD" Then
        Command1.Enabled = True
        Command2.Enabled = False
        Command3.Enabled = False
     ElseIf Combo1.Text = "EDIT" Then
        Command1.Enabled = False
        Command2.Enabled = True
        Command3.Enabled = False
     ElseIf Combo1.Text = "DELETE" Then
        Command1.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = True
        
     End If
    
End Sub


Private Sub Command1_Click()
   'add data to database
    recS.AddNew
    Call textTodata 'to send data to database
    recS.Update
    MsgBox "DONE!!!!"
    
    'text boxes null value
    Call initText
    
    
       
End Sub
Private Sub textTodata()
    'textbox values to database
    recS.Fields(1).Value = Text2.Text
    recS.Fields(2).Value = Text7.Text
    recS.Fields(3).Value = Text3.Text
    recS.Fields(4).Value = Text4.Text
    recS.Fields(5).Value = Text5.Text
   
End Sub
Private Sub initText()
    'initialize with null values
    Text2.Text = "First Name"
    Text7.Text = "Last Name"
    Text3.Text = "Address"
    Text4.Text = "E-mail"
    Text5.Text = "Ph.no"
End Sub
Private Sub Command2_Click()
    
    'recC.EditMode
    Call textTodata
    recS.Update
    MsgBox "DONE!!!!"
    'textbox null
    Call initText

    
End Sub

Private Sub Command3_Click()
recS.Delete
MsgBox "DONE!!!!"
Call initText
End Sub

Private Sub Command4_Click()
    Dim flag As Boolean
    flag = True
    recS.MoveFirst
    While Not recS.EOF = True
            If recS.Fields(1).Value = Text2.Text And recS.Fields(2).Value = Text7.Text Then
                   'initialize text boxes
                    Call init_textboxes
                  'set flag to false
                   flag = False
                   Exit Sub
            End If
            recS.MoveNext
    Wend
    If flag = True Then
        MsgBox ("Invalid Details")
    End If
    
End Sub
Private Sub init_textboxes()
'show data in text boxes
    Text1.Text = recS.Fields(0).Value
    Text2.Text = recS.Fields(1).Value
    Text7.Text = recS.Fields(2).Value
    Text3.Text = recS.Fields(3).Value
    Text4.Text = recS.Fields(4).Value
    Text5.Text = recS.Fields(5).Value
End Sub

Private Sub Form_Load()


frmSupplier.BackColor = RGB(238, 238, 238)
frmSupplier.Top = frmmain.Top + 1000
frmSupplier.Left = frmmain.Left + 3735
frmSupplier.Height = frmmain.Height - 1000
frmSupplier.Width = frmmain.Width - 3735

'**************
Command4.Visible = False 'hide search button
Text1.Enabled = False

'****combo1
 Combo1.AddItem "ADD"
 Combo1.AddItem "EDIT"
 Combo1.AddItem "DELETE"
 
 
 
 'database
 Call Module2.main
 Set recS = New ADODB.Recordset
 recS.Open "select *from Supplier", con, adOpenDynamic, adLockPessimistic, adCmdText
 
 'disable command buttons
 Command1.Enabled = False
 Command2.Enabled = False
 Command3.Enabled = False

End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Call validateA(KeyAscii)
End Sub

Private Sub Text3_Click()
Text3.Text = ""
End Sub
Private Sub Text4_Click()
Text4.Text = ""
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Call validateE(KeyAscii)
End Sub

Private Sub Text5_Click()
Text5.Text = ""
End Sub
Private Sub Text6_Click()
Text6.Text = ""
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
Call validateN(KeyAscii)
End Sub

Private Sub Text7_Click()
Text7.Text = ""
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
Call validateA(KeyAscii)
End Sub
