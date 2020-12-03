VERSION 5.00
Begin VB.Form frmReport 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12570
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame7 
      Height          =   2415
      Left            =   240
      TabIndex        =   18
      Top             =   5520
      Width           =   3855
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   720
         TabIndex        =   21
         Text            =   "Enter Sales ID"
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "SHOW"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   19
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "CUSTOMER BILL"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame6 
      Height          =   2415
      Left            =   8400
      TabIndex        =   15
      Top             =   2880
      Width           =   3855
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   720
         TabIndex        =   22
         Text            =   "Enter Purchase ID"
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton Command6 
         Caption         =   "SHOW"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   16
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "SUPPLIER BILL"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame5 
      Height          =   2415
      Left            =   4320
      TabIndex        =   12
      Top             =   2880
      Width           =   3855
      Begin VB.CommandButton Command5 
         Caption         =   "SHOW"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   13
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "SALES"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2415
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   3855
      Begin VB.CommandButton Command4 
         Caption         =   "SHOW"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   10
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "PURCHASE"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2415
      Left            =   8400
      TabIndex        =   6
      Top             =   240
      Width           =   3855
      Begin VB.CommandButton Command3 
         Caption         =   "SHOW"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   7
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "STOCKS"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   3855
      Begin VB.CommandButton Command2 
         Caption         =   "SHOW"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   4
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "SUPPLIER"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "SHOW"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   2
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "CUSTOMER"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    With customerReport
        .Top = frmmain.Top + 1000
        .Left = frmmain.Left + 3735
        .Height = frmmain.Height - 1000
        .Width = frmmain.Width - 3735
    End With
    customerReport.Refresh
    customerReport.Show
    DataEnvironment1.rsCommand1.Close
End Sub

Private Sub Command2_Click()
    With supplierReport
        .Top = frmmain.Top + 1000
        .Left = frmmain.Left + 3735
        .Height = frmmain.Height - 1000
        .Width = frmmain.Width - 3735
    End With
    supplierReport.Refresh
    supplierReport.Show
    DataEnvironment1.rsCommand2.Close
End Sub

Private Sub Command3_Click()
    With stocksReport
        .Top = frmmain.Top + 1000
        .Left = frmmain.Left + 3735
        .Height = frmmain.Height - 1000
        .Width = frmmain.Width - 3735
    End With
    stocksReport.Refresh
    stocksReport.Show
    DataEnvironment1.rsCommand3.Close
End Sub

Private Sub Command6_Click()
    With supplierBillReport
        .Top = frmmain.Top + 1000
        .Left = frmmain.Left + 3735
        .Height = frmmain.Height - 1000
        .Width = frmmain.Width - 3735
    End With
    DataEnvironment1.supplierBillcmd Text2
    supplierBillReport.Show
    supplierBillReport.Refresh
    DataEnvironment1.rssupplierBillcmd.Close
End Sub

Private Sub Command7_Click()
    With customerBillReport
        .Top = frmmain.Top + 1000
        .Left = frmmain.Left + 3735
        .Height = frmmain.Height - 1000
        .Width = frmmain.Width - 3735
    End With
    DataEnvironment1.customerBillcmd Text1
    customerBillReport.Show
    customerBillReport.Refresh
    DataEnvironment1.rscustomerBillcmd.Close
End Sub

Private Sub Form_Load()
'***form location
With frmReport
    .BackColor = RGB(238, 238, 238)
    .Top = frmmain.Top + 1000
    .Left = frmmain.Left + 3735
    .Height = frmmain.Height - 1000
    .Width = frmmain.Width - 3735
End With
'++++++++++++++++++++++++++++++++++
End Sub

