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
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
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
      Left            =   3960
      TabIndex        =   13
      Top             =   6360
      Width           =   4575
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   6120
      TabIndex        =   11
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6120
      TabIndex        =   7
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Ph no"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Details"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   12015
   End
End
Attribute VB_Name = "frmcustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmcustomer.BackColor = RGB(240, 150, 118)
frmcustomer.Top = frmmain.Top
frmcustomer.Left = frmmain.Left + 3735
frmcustomer.Height = frmmain.Height
frmcustomer.Width = frmmain.Width - 3735


End Sub
