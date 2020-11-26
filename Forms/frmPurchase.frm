VERSION 5.00
Begin VB.Form frmPurchase 
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
   Begin VB.CommandButton Command7 
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
      Left            =   8520
      TabIndex        =   14
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SEARCH"
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
      Left            =   5280
      TabIndex        =   13
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SEARCH"
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
      Left            =   7680
      TabIndex        =   12
      Top             =   2640
      Width           =   1935
   End
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
      Left            =   2880
      TabIndex        =   2
      Text            =   "Hardware Name"
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Text            =   "Supplier First Name"
      Top             =   2640
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
      Left            =   7680
      TabIndex        =   5
      Text            =   "QTY"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Text            =   "Ref.ID"
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Text            =   "Supplier Last Name"
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Text            =   "P-ID"
      Top             =   1800
      Width           =   1935
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recP, recS, recIn, recSt As ADODB.Recordset
Dim tamt As Integer
Private Sub Combo1_click()
    If Combo1.Text = "EDIT" Or Combo1.Text = "DELETE" Then
        Combo2.Visible = True
        Command4.Visible = True
    Else
        Combo2.Visible = False
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
    recP.AddNew
    Call textTodata 'to send data to database
    recP.Update
    MsgBox "DONE!!!!"
    
    'text boxes null value
    Call initText
    
    
       
End Sub
Private Sub textTodata()
    'textbox values to database
    recP.Fields(1).Value = recS.Fields(0).Value
    recP.Fields(2).Value = Text4.Text
    recP.Fields(3).Value = tamt
    recP.Fields(4).Value = Format(Now, "mm/dd/yy hh:mm:ss")
    recP.Fields(5).Value = 1
   
End Sub
Private Sub initText()
    'initialize with null values
    Text2.Text = "Supplier First Name"
    Text7.Text = "Hardware Name"
    Text3.Text = "Supplier Last Name"
    Text4.Text = "Ref.ID"
    Text5.Text = "QTY"
End Sub
Private Sub Command2_Click()
    
    'recC.EditMode
    Call textTodata
    recP.Update
    MsgBox "DONE!!!!"
    'textbox null
    Call initText

    
End Sub

Private Sub Command3_Click()
recP.Delete
MsgBox "DONE!!!!"
Call initText
End Sub

Private Sub Command4_Click()
    Dim flag As Boolean
    If Combo2.Text = "SID" Then
        flag = True
        recP.MoveFirst
        While Not recP.EOF = True
            If recP.Fields(0).Value = Text1.Text Then
                 'show data in text boxes
                  Call init_textboxes
                  'set flag to false
                  flag = False
                  Exit Sub
            End If
            recP.MoveNext
        Wend
        If flag = True Then
            MsgBox ("Invalid Details")
        End If
     ElseIf Combo2.Text = "NAME" Then
        flag = True
        recP.MoveFirst
        While Not recP.EOF = True
            If recP.Fields(1).Value = Text2.Text And recP.Fields(2).Value = Text7.Text Then
                   'initialize text boxes
                    Call init_textboxes
                  'set flag to false
                   flag = False
                   Exit Sub
            End If
            recP.MoveNext
        Wend
        If flag = True Then
            MsgBox ("Invalid Details")
        End If
     End If
    
End Sub
Private Sub init_textboxes()
'show data in text boxes
    Text2.Text = recP.Fields(1).Value
    Text7.Text = recP.Fields(2).Value
    Text3.Text = recP.Fields(3).Value
    Text4.Text = recP.Fields(4).Value
    Text5.Text = recP.Fields(5).Value
End Sub

Private Sub Command5_Click()
    flag = True
        recS.MoveFirst
        While Not recS.EOF = True
            If recS.Fields(1).Value = Text2.Text And recS.Fields(2).Value = Text3.Text Then
                   'initialize text boxes
                    Call init_textboxesSup
                    MsgBox "Supplier ID : " & recS.Fields(0).Value
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
Private Sub init_textboxesSup()
    Text2.Text = recS.Fields(1).Value
    Text2.Text = recS.Fields(1).Value
End Sub

Private Sub Command6_Click()
    flag = True
        recSt.MoveFirst
        While Not recSt.EOF = True
            If recSt.Fields(2).Value = Text7.Text Then
                   'initialize text boxes
                    Call init_textboxesHard
                    MsgBox "Found"
                  'set flag to false
                   flag = False
                   Exit Sub
            End If
            recSt.MoveNext
        Wend
        If flag = True Then
            MsgBox ("Invalid Details")
        End If
End Sub
Private Sub init_textboxesHard()
    Text7.Text = recSt.Fields(2).Value

End Sub

Private Sub Command7_Click()
    If (recSt.Fields(5) - CInt(Text5.Text)) < 0 Then
        MsgBox "Only " & recSt.Fields(5).Value & " is available."
    Else
        recIn.AddNew
        recIn.Fields(1).Value = Text4.Text
        recIn.Fields(2).Value = recSt.Fields(0).Value
        recIn.Fields(3).Value = Text5.Text
        'total amt
        tamt = tamt + CInt(Text5.Text) * recSt.Fields(4).Value
        MsgBox tamt
        recSt.Fields(5).Value = recSt.Fields(5).Value - CInt(Text5.Text)
        MsgBox "Item Added"
        recIn.Update
    End If
    
End Sub

Private Sub Form_Load()

'*****amt
tamt = 0

'***form location
frmPurchase.BackColor = RGB(238, 238, 238)
frmPurchase.Top = frmmain.Top + 1000
frmPurchase.Left = frmmain.Left + 3735
frmPurchase.Height = frmmain.Height - 1000
frmPurchase.Width = frmmain.Width - 3735




'****combo1
 Combo1.AddItem "ADD"
 Combo1.AddItem "EDIT"
 Combo1.AddItem "DELETE"
 
 '***combo2
 Combo2.AddItem "SID"
 Combo2.AddItem "NAME"
 Combo2.Visible = False
 Command4.Visible = False
 
 'database
 Call Module2.main
 Set recP = New ADODB.Recordset
 recP.Open "select *from Purchase", con, adOpenDynamic, adLockPessimistic, adCmdText
    'supplier
 Set recS = New ADODB.Recordset
 recS.Open "select *from Supplier", con, adOpenDynamic, adLockPessimistic, adCmdText
    'IN
 Set recIn = New ADODB.Recordset
 recIn.Open "select *from Stock_in", con, adOpenDynamic, adLockPessimistic, adCmdText
    'Hardware
 Set recSt = New ADODB.Recordset
 recSt.Open "select *from Hardware", con, adOpenDynamic, adLockPessimistic, adCmdText
 
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
