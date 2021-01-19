VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSales 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Sales"
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12570
   LinkTopic       =   "Form2"
   ScaleHeight     =   8205
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "PAY"
      Height          =   615
      Left            =   6240
      TabIndex        =   12
      Top             =   5880
      Width           =   1695
   End
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
      RecordSource    =   "select * from Stock_out"
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
      TabIndex        =   10
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   615
      Left            =   8400
      TabIndex        =   9
      Top             =   5880
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   3495
      Left            =   3960
      TabIndex        =   8
      Top             =   2040
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
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
         TabIndex        =   11
         Top             =   240
         Width           =   2055
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
         Caption         =   "CUSTOMER "
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   12
            Charset         =   0
            Weight          =   600
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
         Caption         =   "S-ID"
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
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
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
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public cid, cusid, datagrid1_id, datagrid2_id As Integer
Dim rec As New ADODB.Recordset
Dim auto, recStdIn As ADODB.Recordset


Private Sub Combo1_click()
If Combo1.Text <> "" Then
    Adodc1.RecordSource = "select * from Hardware where type='" & Combo1 & "'"
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    
    DataGrid1.Refresh
End If
End Sub




Private Sub Combo3_Click()
    Adodc1.RecordSource = "select * from Customer where fname='" & Combo2 & "' and lname='" & Combo3 & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount = 0 Then
        MsgBox "Invalid Customer Data", vbExclamation
    Else
        cusid = Adodc1.Recordset.Fields(0).Value
        DataGrid1.Refresh
    End If
    

End Sub

Private Sub Command1_Click()

Dim tempid As Integer
tempid = Val(Text2.Text)
'""""""""""""""""""""""""""""""" duplicate data entry checker
Dim recStdcheck As ADODB.Recordset
Set recStdcheck = New ADODB.Recordset
recStdcheck.Open "Select Hardware.ID From Hardware,Stock_out Where Hardware.Id=Stock_out.hid And Stock_out.outid=" & Text2.Text & " ", con, adOpenDynamic, adLockPessimistic, adCmdText

If recStdcheck.RecordCount > 0 Then
   recStdcheck.MoveFirst
End If
While Not (recStdcheck.EOF)
    'Print recStdcheck.Fields(0).Value
    'Print datagrid1_id
    Dim a, b As Integer
    a = Val(recStdcheck.Fields(0).Value)
    b = Val(datagrid1_id)
    If a = b Then
        MsgBox "Duplicate Data", vbCritical
        Exit Sub
    End If
    recStdcheck.MoveNext
Wend
'""""""""""""""""""""""""""""""""""

With rec
      .AddNew
      .Fields(1).Value = Val(Text2.Text)
      .Fields(2).Value = datagrid1_id
      .Fields(3).Value = CInt(InputBox("Enter Quentity"))
      .Update
End With

'""""""""""""""""""""""""""" datagrid 2
'Dim recStdIn As ADODB.Recordset
Set recStdIn = New ADODB.Recordset
recStdIn.Open "Select Hardware.ID,hname,desc From Hardware,Stock_out Where Hardware.Id=Stock_out.hid And Stock_out.outid=" & Text2.Text & " ", con, adOpenDynamic, adLockPessimistic, adCmdText
'datagrid2
Set DataGrid2.DataSource = recStdIn
DataGrid2.Refresh
'"""""""""""""""""""""""""""""""
End Sub

Private Sub Command2_Click()
 
 Print auto.RecordCount
 'If ((Not (auto.EOF)) Or (Not (auto.BOF))) Then
  If auto.RecordCount > 0 Then
     If auto.RecordCount > 1 Then
        While Not (auto.RecordCount = 0)
          auto.MoveLast
          auto.Delete
        Wend
      Else
       auto.Delete
      End If
  Else
     MsgBox "Not selected"
  End If
 'End If
 '""""""""""""""""""""""""""" datagrid 2
 'Dim recStdIn As ADODB.Recordset
 Set recStdIn = New ADODB.Recordset
 recStdIn.Open "Select Hardware.ID,hname,desc From Hardware,Stock_out Where Hardware.Id=Stock_out.hid And Stock_out.outid=" & Text2.Text & " ", con, adOpenDynamic, adLockPessimistic, adCmdText
 'datagrid2
 Set DataGrid2.DataSource = recStdIn
 DataGrid2.Refresh
'"""""""""""""""""""""""""""""""


End Sub

Private Sub Command3_Click()
Dim amt As Long
'"""""""""""""""""""""""" connection == Stock_out
Dim temp As ADODB.Recordset
Set temp = New ADODB.Recordset
temp.Open "select *from Stock_out where outid = " & Text2.Text & "", con, adOpenDynamic, adLockPessimistic, adCmdText
'===================================

'"""""""""""""""""""" connection == Hardware
Dim temphard As ADODB.Recordset
Set temphard = New ADODB.Recordset
temphard.Open "select *from Hardware", con, adOpenDynamic, adLockPessimistic, adCmdText
'==================================

amt = 0
    temp.MoveFirst 'stock out first data
    While Not (temp.EOF) 'stock out iterator
        Dim tempidh As Integer
        tempidh = temp.Fields(2).Value
        
        temphard.MoveFirst  'hardware first data
        While Not (temphard.EOF)
            If (temphard.Fields(0).Value = temp.Fields(2).Value) Then
               amt = amt + ((temphard.Fields(6).Value) * (temp.Fields(3).Value))
               'MsgBox amt
               temphard.Fields(5).Value = temphard.Fields(5).Value - temp.Fields(3).Value
            End If
            temphard.MoveNext
        Wend
        temp.MoveNext
    Wend
MsgBox "Total payable amount : " & amt



'''''''''''''''insert data into Sales table
    Dim purc As ADODB.Recordset
    Set purc = New ADODB.Recordset
    purc.Open "select *from Sales", con, adOpenDynamic, adLockPessimistic, adCmdText
    With purc
        .AddNew
        .Fields(0).Value = CInt(Text2.Text)
        .Fields(1).Value = cusid
        .Fields(2).Value = CInt(Text2.Text)
        .Fields(3).Value = amt
        .Fields(4).Value = Date
        .Fields(5).Value = 1 'i am on main form thats why default user id is 1
        .Update
    End With
    MsgBox "Transaction Successful!!!!"
'''''''''************



'bill report
    With customerBillReport
        .Top = frmmain.Top + 1000
        .Left = frmmain.Left + 3735
        .Height = frmmain.Height - 1000
        .Width = frmmain.Width - 3735
    End With
    '+++++++++++++++++++++++++++ customer name - dynamic
    Dim namePrinter As ADODB.Recordset
    Set namePrinter = New ADODB.Recordset
    namePrinter.Open "select fname,lname,amt from Customer,Sales where Sales.ID = " & Text2.Text & " AND Customer.ID=cid", con, adOpenDynamic, adLockPessimistic, adCmdText
    customerBillReport.Sections.Item("Section4").Controls.Item("Label4").Caption = namePrinter.Fields("fname").Value
    customerBillReport.Sections.Item("Section4").Controls.Item("Label12").Caption = namePrinter.Fields("lname").Value
    customerBillReport.Sections.Item("Section5").Controls.Item("Label7").Caption = "Total amount : " & namePrinter.Fields("amt").Value & " Rs."
    '++++++++++++++++++++++++++++++++++++++++++++++++
DataEnvironment1.customerBillcmd Text2
customerBillReport.Refresh
customerBillReport.Show
DataEnvironment1.rscustomerBillcmd.Close
''''''''''''''********************


'''''''''''clear form and  generat next ID
Call Auto_num
Combo2.Text = "First Name"
Combo3.Text = "Last Name"
Combo1.Text = "Processor"
Set DataGrid1.DataSource = Nothing
Set DataGrid2.DataSource = Nothing

'''''''''''''''''''''*********************

End Sub

Private Sub DataGrid1_Click()
If DataGrid1.DataSource Is Nothing Then
    MsgBox "Empty"
Else
   'Print DataGrid1.Columns(0).Value
   datagrid1_id = DataGrid1.Columns(0).Value
End If
End Sub

Private Sub DataGrid2_Click()
If DataGrid2.DataSource Is Nothing Then
    MsgBox "Empty"
Else
   'Print DataGrid2.Columns(0).Value
   datagrid2_id = DataGrid2.Columns(0).Value
   Set auto = New ADODB.Recordset
   auto.Open "select *from Stock_out where outid=" & Text2.Text & " and hid =" & DataGrid2.Columns(0).Value & "", con, adOpenDynamic, adLockPessimistic, adCmdText
 
End If
End Sub

Private Sub Form_Load()
'************************adodc
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Databases\ProData.mdb;Persist Security Info=False"
S.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Databases\ProData.mdb;Persist Security Info=False"

'***form location
With frmSales
.BackColor = RGB(238, 238, 238)
.Top = frmmain.Top + 1000
.Left = frmmain.Left + 3735
.Height = frmmain.Height - 1000
.Width = frmmain.Width - 3735
End With
'******adodc1 and adding fname,lname to combo boxes
Adodc1.RecordSource = "select * from Customer"
Adodc1.Refresh

With Adodc1.Recordset
Do Until .EOF
Combo2.AddItem ![fname]
Combo3.AddItem ![lname]
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
rec.Open "select *from Stock_out", con, adOpenDynamic, adLockPessimistic, adCmdText


'"""""""""""""""""""""""""""""""""""" S and datagrid2

''''''''''
'"""""""""""""""""""""""""""""""""""


End Sub
Private Sub Auto_num()


  '  'displaying of the Auto P ID
   ' Order By Product_Name DESC
   Call Module2.main
   Set auto = New ADODB.Recordset
   auto.Open "select *from Sales", con, adOpenDynamic, adLockPessimistic, adCmdText
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

