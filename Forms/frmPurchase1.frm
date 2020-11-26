VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
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
   Begin MSAdodcLib.Adodc S 
      Height          =   375
      Left            =   8400
      Top             =   7440
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\HYDRA\Desktop\Sem-5-pro\Databases\ProData.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\HYDRA\Desktop\Sem-5-pro\Databases\ProData.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Stock_in"
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton Command2 
      Caption         =   "REMOVE"
      Height          =   615
      Left            =   10560
      TabIndex        =   11
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   615
      Left            =   8400
      TabIndex        =   10
      Top             =   5880
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmPurchase1.frx":0000
      Height          =   3495
      Left            =   3960
      TabIndex        =   9
      Top             =   2040
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6165
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
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   3495
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Text            =   "Select Hardware Type"
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12255
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   2160
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   9840
         TabIndex        =   8
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   91881473
         CurrentDate     =   44156
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   4320
         TabIndex        =   5
         Text            =   "Last Name"
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2160
         TabIndex        =   3
         Text            =   "First Name"
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "SUPPLIER :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "P-ID"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmPurchase1.frx":0010
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   8281
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5520
      Top             =   7440
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\HYDRA\Desktop\Sem-5-pro\Databases\ProData.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\HYDRA\Desktop\Sem-5-pro\Databases\ProData.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from Hardware"
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
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sid As Integer
Dim rec As New ADODB.Recordset
Dim auto As ADODB.Recordset


Private Sub Combo1_click()
If Combo1.Text <> "" Then
    Adodc1.RecordSource = "select * from Hardware where type='" & Combo1 & "'"
    Adodc1.Refresh
    DataGrid1.Refresh
End If
End Sub

Private Sub Combo2_Click()
If Combo2.Text <> "" Then
    Adodc1.RecordSource = "select * from Supplier where fname='" & Combo2 & "'"
    Adodc1.Refresh
    DataGrid1.Refresh
    With Adodc1.Recordset
            Do Until .EOF
            Combo3.AddItem ![lname]
            .MoveNext
            Loop
    End With
End If
End Sub

Private Sub Combo3_Click()
    Adodc1.RecordSource = "select * from Supplier where fname='" & Combo2 & "' and lname='" & Combo3 & "'"
    Adodc1.Refresh
    DataGrid1.Refresh
    Text1.Text = Adodc1.Recordset.Fields(0).Value
    

End Sub

Private Sub Command1_Click()
Dim tempid As Integer
tempid = Val(Text2.Text)
With rec
      .AddNew
      .Fields(1).Value = Val(Text2.Text)
      .Fields(2).Value = DataGrid1.Columns(0).Value
      .Fields(3).Value = CInt(InputBox("Enter Quentity"))
      .Update
End With
    S.RecordSource = "Select Hardware.ID,hname,desc From Hardware,Stock_in Where Hardware.Id=Stock_in.hid And Stock_in.inid=" & Text2.Text & " "
    S.Refresh
    DataGrid2.Refresh
End Sub

Private Sub Command2_Click()
 Set auto = New ADODB.Recordset
 auto.Open "select *from Stock_in where inid=" & Text2.Text & " and hid=" & DataGrid1.Columns(0).Value & "", con, adOpenDynamic, adLockPessimistic, adCmdText
 If ((Not (auto.EOF)) Or (Not (auto.BOF))) Then
     auto.Delete
 End If
 S.Refresh
 DataGrid1.Refresh

End Sub

Private Sub Form_Load()


'***form location
With Form2
.BackColor = RGB(238, 238, 238)
.Top = frmmain.Top + 1000
.Left = frmmain.Left + 3735
.Height = frmmain.Height - 1000
.Width = frmmain.Width - 3735
End With
'******adodc1
Adodc1.RecordSource = "select * from Supplier"
Adodc1.Refresh

With Adodc1.Recordset
Do Until .EOF
Combo2.AddItem ![fname]

.MoveNext
Loop
End With

'******
Call Auto_num



'*******combo1
With Combo1
    .AddItem "Processor"
    .AddItem "Motherboard"
    .AddItem "PSU"
    .AddItem "RAM"
    .AddItem "Hard Disk"
    .AddItem "Graphics Card"
End With

'*****database
Call Module2.main
Set rec = New ADODB.Recordset
rec.Open "select *from Stock_in", con, adOpenDynamic, adLockPessimistic, adCmdText


End Sub
Private Sub Auto_num()


  '  'displaying of the Auto P ID
   ' Order By Product_Name DESC
   Call Module2.main
   Set auto = New ADODB.Recordset
   auto.Open "select *from Purchase", con, adOpenDynamic, adLockPessimistic, adCmdText
    With auto
        If .RecordCount = 0 Then
           Text2.Text = 1
        Else
           .MoveLast
           Text2.Text = (auto!id) + 1
        End If
        
        auto.Close
        Set auto = Nothing
        End With
        Text2.Locked = True

End Sub

