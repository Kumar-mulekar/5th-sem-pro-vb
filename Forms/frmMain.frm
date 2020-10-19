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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9015
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "@Arial Unicode MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   3735
      Begin XtremeSkinFramework.SkinFramework skn 
         Left            =   480
         Top             =   1800
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
'**********frmMain**********
skn.LoadSkin App.Path + "\Styles\Office2010.cjstyles", ""
'Office2007 Office2010  WinXP.Luna WinXP.Royale Codejock
skn.ApplyWindow Me.hWnd

frmmain.BackColor = RGB(238, 238, 238)



'**************frame 1***************
Frame1.BackColor = RGB(51, 49, 44)
Label1.BackColor = RGB(51, 49, 44)
Label2.BackColor = RGB(51, 49, 44)
Label3.BackColor = RGB(51, 49, 44)
Label4.BackColor = RGB(51, 49, 44)
Label5.BackColor = RGB(51, 49, 44)
Label6.BackColor = RGB(51, 49, 44)
Label7.BackColor = RGB(51, 49, 44)
Label8.BackColor = RGB(51, 49, 44)



'*************frame 2************
Frame2.Left = frmmain.Left
Frame2.Width = frmmain.Width
Frame2.Top = frmmain.Top
Frame2.Height = 1000
Frame2.BackColor = RGB(255, 255, 255)


'*********label9*******
Label9.Left = Frame2.Left
Label9.Top = Frame2.Top
Label9.Height = Frame2.Height
Label9.Width = Frame1.Width
Label9.BackColor = RGB(51, 49, 44)
Label9.Caption = "Admin" '"USER:" + userName

End Sub

Private Sub Label1_Click()
lblpre = lblcl
lblcl = 1
Call lblclick
Label1.BackColor = RGB(0, 172, 193)

End Sub

Private Sub Label2_Click()
lblpre = lblcl
lblcl = 2
Call lblclick
Label2.BackColor = RGB(0, 172, 193)
frmcustomer.Show
End Sub

Private Sub Label3_Click()
lblpre = lblcl
lblcl = 3
Call lblclick
Label3.BackColor = RGB(0, 172, 193)
End Sub

Private Sub Label4_Click()
lblpre = lblcl
lblcl = 4
Call lblclick
Label4.BackColor = RGB(0, 172, 193)
End Sub

Private Sub Label5_Click()
lblpre = lblcl
lblcl = 5
Call lblclick
Label5.BackColor = RGB(0, 172, 193)
End Sub

Private Sub Label6_Click()
lblpre = lblcl
lblcl = 6
Call lblclick
Label6.BackColor = RGB(0, 172, 193)
End Sub

Private Sub Label7_Click()
lblpre = lblcl
lblcl = 7
Call lblclick
Label7.BackColor = RGB(0, 172, 193)
End Sub

Private Sub Label8_Click()
End
End Sub
Public Sub lblclick()
Select Case (lblpre)
   Case 1:
   Label1.BackColor = RGB(51, 49, 44)
   Case 2:
Label2.BackColor = RGB(51, 49, 44)
Case 3:
Label3.BackColor = RGB(51, 49, 44)
Case 4:
Label4.BackColor = RGB(51, 49, 44)
Case 5:
Label5.BackColor = RGB(51, 49, 44)
Case 6:
Label6.BackColor = RGB(51, 49, 44)
Case 7:
Label7.BackColor = RGB(51, 49, 44)
Case 8:
Label8.BackColor = RGB(51, 49, 44)
End Select

         

End Sub
