VERSION 5.00
Begin VB.Form frmcustomer 
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
      TabIndex        =   11
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5280
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
      Text            =   "C-ID"
      Top             =   2520
      Width           =   1935
   End
End
Attribute VB_Name = "frmcustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recC As ADODB.Recordset
Private Sub Combo1_click()
    If Combo1.Text = "EDIT" Or Combo1.Text = "DELETE" Then
        Combo2.Visible = True
        Command4.Visible = True
    Else
        Combo2.Visible = False
        Command4.Visible = False
    End If
    
End Sub


Private Sub Command1_Click()
   'add data to database
    recC.AddNew
    'recC.Fields(0).Value = Text1.Text
    recC.Fields(1).Value = Text2.Text
    recC.Fields(2).Value = Text7.Text
    recC.Fields(3).Value = Text3.Text
    recC.Fields(4).Value = Text4.Text
    recC.Fields(5).Value = Text5.Text
    recC.Update
    MsgBox "DONE!!!!"
    
    'text boxes
    Text2.Text = "First Name"
    Text7.Text = "Last Name"
    Text3.Text = "Address"
    Text4.Text = "E-mail"
    Text5.Text = "Ph.no"
   
    
    
End Sub

Private Sub Command4_Click()
    Dim flag As Boolean
    If Combo2.Text = "CID" Then
        flag = True
        recC.MoveFirst
        While Not recC.EOF = True
            If recC.Fields(0).Value = Text1.Text Then
                 'show data in text boxes
                  Call init_textboxes
                  'set flag to false
                  flag = False
            End If
            recC.MoveNext
        Wend
        If flag = True Then
            MsgBox ("Invalid Details")
        End If
     ElseIf Combo2.Text = "NAME" Then
        flag = True
        recC.MoveFirst
        While Not recC.EOF = True
            If recC.Fields(1).Value = Text2.Text And recC.Fields(2).Value = Text7.Text Then
                   'initialize text boxes
                    Call init_textboxes
                  'set flag to false
                   flag = False
            End If
            recC.MoveNext
        Wend
        If flag = True Then
            MsgBox ("Invalid Details")
        End If
     End If
    
End Sub
Private Sub init_textboxes()
'show data in text boxes
    Text2.Text = recC.Fields(1).Value
    Text7.Text = recC.Fields(2).Value
    Text3.Text = recC.Fields(3).Value
    Text4.Text = recC.Fields(4).Value
End Sub

Private Sub Form_Load()


frmcustomer.BackColor = RGB(238, 238, 238)
frmcustomer.Top = frmmain.Top + 1000
frmcustomer.Left = frmmain.Left + 3735
frmcustomer.Height = frmmain.Height - 1000
frmcustomer.Width = frmmain.Width - 3735

'********shapes cust******


'****combo1
 Combo1.AddItem "ADD"
 Combo1.AddItem "EDIT"
 Combo1.AddItem "DELETE"
 
 '***combo2
 Combo2.AddItem "CID"
 Combo2.AddItem "NAME"
 Combo2.Visible = False
 Command4.Visible = False
 
 'database
 Call Module2.main
 Set recC = New ADODB.Recordset
 recC.Open "select *from Customer", con, adOpenDynamic, adLockPessimistic, adCmdText


End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub
Private Sub Text3_Click()
Text3.Text = ""
End Sub
Private Sub Text4_Click()
Text4.Text = ""
End Sub
Private Sub Text5_Click()
Text5.Text = ""
End Sub
Private Sub Text6_Click()
Text6.Text = ""
End Sub
Private Sub Text7_Click()
Text7.Text = ""
End Sub
