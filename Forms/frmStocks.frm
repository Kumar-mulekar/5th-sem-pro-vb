VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStocks 
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5415
      Left            =   2400
      TabIndex        =   12
      Top             =   1200
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9551
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show Stocks"
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
      TabIndex        =   11
      Top             =   240
      Width           =   6735
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2880
      TabIndex        =   10
      Text            =   "TYPE"
      Top             =   2520
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   6000
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2880
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Text            =   "Name"
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
      TabIndex        =   5
      Top             =   6000
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Text            =   "QTY"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Text            =   "Price"
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Text            =   "Description"
      Top             =   3360
      Width           =   6735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Text            =   "H-ID"
      Top             =   2520
      Width           =   1935
   End
End
Attribute VB_Name = "frmStocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recSt As ADODB.Recordset
Private Sub Combo1_click()
    If Combo1.Text = "EDIT" Or Combo1.Text = "DELETE" Then
        MsgBox "Enter Hardware Name then tap SEARCH"
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
    recSt.AddNew
    Call textTodata 'to send data to database
    recSt.Update
    MsgBox "DONE!!!!"
    
    'text boxes null value
    Call initText
    
    
       
End Sub
Private Sub textTodata()
    'textbox values to database
    recSt.Fields(1).Value = Combo3.Text
    recSt.Fields(2).Value = Text2.Text
    recSt.Fields(3).Value = Text3.Text
    recSt.Fields(4).Value = Text4.Text
    recSt.Fields(5).Value = Text5.Text
    recSt.Fields(6).Value = Val(Text4.Text) + (Val(Text4.Text) \ 10)
End Sub
Private Sub initText()
    'initialize with null values
    Combo3.Text = "TYPE"
    Text2.Text = "Name"
    Text3.Text = "Description"
    Text4.Text = "Price"
    Text5.Text = "QTY"
End Sub
Private Sub Command2_Click()
    
    'recC.EditMode
    Call textTodata
    recSt.Update
    MsgBox "DONE!!!!"
    'textbox null
    Call initText

    
End Sub

Private Sub Command3_Click()
recSt.Delete
MsgBox "DONE!!!!"
Call initText
End Sub

Private Sub Command4_Click()
    Dim flag As Boolean
    flag = True
    recSt.MoveFirst
    While Not recSt.EOF = True
            If recSt.Fields(2).Value = Text2.Text Then
                   'initialize text boxes
                    Call init_textboxes
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
Private Sub init_textboxes()
'show data in text boxes
    Text1.Text = recSt.Fields(0).Value
    Combo3.Text = recSt.Fields(1).Value
    Text2.Text = recSt.Fields(2).Value
    Text3.Text = recSt.Fields(3).Value
    Text4.Text = recSt.Fields(4).Value
    Text5.Text = recSt.Fields(5).Value
End Sub

Private Sub Command5_Click()
    If Command5.Caption = "Show Stocks" Then
        Command5.Caption = "Tap here to enter data"
        DataGrid1.Visible = True
        DataGrid1.Refresh
    ElseIf Command5.Caption = "Tap here to enter data" Then
        Command5.Caption = "Show Stocks"
        DataGrid1.Visible = False
    End If
        
End Sub

Private Sub Form_Load()


frmStocks.BackColor = RGB(238, 238, 238)
frmStocks.Top = frmmain.Top + 1000
frmStocks.Left = frmmain.Left + 3735
frmStocks.Height = frmmain.Height - 1000
frmStocks.Width = frmmain.Width - 3735

'**************
Command4.Visible = False 'hide search button
Text1.Enabled = False


'****combo1
 Combo1.AddItem "ADD"
 Combo1.AddItem "EDIT"
 Combo1.AddItem "DELETE"
 
 
 
 
 '**********combo3
 Combo3.AddItem "Processor"
 Combo3.AddItem "RAM"
 Combo3.AddItem "Motherboard"
 Combo3.AddItem "Graphics Card"
 Combo3.AddItem "PSU"
 Combo3.AddItem "Cabinet"
 Combo3.AddItem "Cooler"
 
 'database
 Call Module2.main
 Set recSt = New ADODB.Recordset
 recSt.Open "select *from Hardware", con, adOpenDynamic, adLockPessimistic, adCmdText
 
 '*******data grid
 Set DataGrid1.DataSource = recSt
 DataGrid1.Refresh
 DataGrid1.Visible = False
 
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
Call validateAN(KeyAscii)
End Sub

Private Sub Text3_Click()
Text3.Text = ""
End Sub

Private Sub Text4_Click()
Text4.Text = ""
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Call validateN(KeyAscii)
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

Private Sub Text5_KeyPress(KeyAscii As Integer)
Call validateN(KeyAscii)
End Sub
