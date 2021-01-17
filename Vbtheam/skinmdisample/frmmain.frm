VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#18.6#0"; "CO23AE~1.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17010
   LinkTopic       =   "Form2"
   ScaleHeight     =   9345
   ScaleWidth      =   17010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   3735
      Begin VB.Line Line1 
         DrawMode        =   1  'Blackness
         X1              =   3720
         X2              =   3720
         Y1              =   120
         Y2              =   9480
      End
      Begin XtremeSkinFramework.SkinFramework skn 
         Left            =   960
         Top             =   840
         _Version        =   1179654
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
      Begin VB.Label Label8 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   8
         Top             =   7080
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "REPORTS"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   7
         Top             =   6480
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "STOCKS"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   6
         Top             =   5880
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "PURCHASE"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   5
         Top             =   5280
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "SALES"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   4
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "SUPPLIER"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   3
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "CUSTOMER"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   2
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "HOME"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   1
         Top             =   2880
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lblcl, lblpre As Integer

Private Sub Form_Load()
lblcl = 0
skn.LoadSkin App.Path + "\Styles\Office2010.cjstyles", ""
'Office2007 Office2010  WinXP.Luna WinXP.Royale Codejock
skn.ApplyWindow Me.hWnd
Frame1.BackColor = RGB(121, 239, 208)
Label1.BackColor = RGB(121, 239, 208)
Label2.BackColor = RGB(121, 239, 208)
Label3.BackColor = RGB(121, 239, 208)
Label4.BackColor = RGB(121, 239, 208)
Label5.BackColor = RGB(121, 239, 208)
Label6.BackColor = RGB(121, 239, 208)
Label7.BackColor = RGB(121, 239, 208)
Label8.BackColor = RGB(121, 239, 208)
End Sub

Private Sub Label1_Click()
lblpre = lblcl
lblcl = 1
Call lblclick
Label1.BackColor = RGB(240, 150, 118)
End Sub

Private Sub Label2_Click()
lblpre = lblcl
lblcl = 2
Call lblclick
Label2.BackColor = RGB(240, 150, 118)
frmcustomer.Show
End Sub

Private Sub Label3_Click()
lblpre = lblcl
lblcl = 3
Call lblclick
Label3.BackColor = RGB(240, 150, 118)
End Sub

Private Sub Label4_Click()
lblpre = lblcl
lblcl = 4
Call lblclick
Label4.BackColor = RGB(240, 150, 118)
End Sub

Private Sub Label5_Click()
lblpre = lblcl
lblcl = 5
Call lblclick
Label5.BackColor = RGB(240, 150, 118)
End Sub

Private Sub Label6_Click()
lblpre = lblcl
lblcl = 6
Call lblclick
Label6.BackColor = RGB(240, 150, 118)
End Sub

Private Sub Label7_Click()
lblpre = lblcl
lblcl = 7
Call lblclick
Label7.BackColor = RGB(240, 150, 118)
End Sub

Private Sub Label8_Click()
End
End Sub
Public Sub lblclick()
Select Case (lblpre)
   Case 1:
   Label1.BackColor = RGB(121, 239, 208)
   Case 2:
Label2.BackColor = RGB(121, 239, 208)
Case 3:
Label3.BackColor = RGB(121, 239, 208)
Case 4:
Label4.BackColor = RGB(121, 239, 208)
Case 5:
Label5.BackColor = RGB(121, 239, 208)
Case 6:
Label6.BackColor = RGB(121, 239, 208)
Case 7:
Label7.BackColor = RGB(121, 239, 208)
Case 8:
Label8.BackColor = RGB(121, 239, 208)
End Select

         

End Sub
