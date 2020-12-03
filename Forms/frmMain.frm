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
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11655
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   3960
         TabIndex        =   11
         Top             =   120
         Width           =   6735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   20.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
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
      Begin VB.Label Label11 
         Caption         =   "USER"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   12
         Top             =   6000
         Width           =   2415
      End
      Begin VB.Line Line1 
         DrawMode        =   8  'Xor Pen
         X1              =   0
         X2              =   3720
         Y1              =   1200
         Y2              =   1200
      End
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
            Name            =   "Lucida Bright"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
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
            Name            =   "Lucida Bright"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   7
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Label Label6 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   6
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   5
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   4
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   3
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   2
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "HOME"
         BeginProperty Font 
            Name            =   "Lucida Bright"
            Size            =   15.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   1
         Top             =   1800
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
Label11.BackColor = RGB(51, 49, 44)


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
Label9.Caption = userName

'+++++++++++++++++++++ user access and hiding user lbl from user
If userAccess = "Admin" Then
    Label11.Visible = True
Else
    Label11.Visible = False
End If

'++++++++++++++++++++++++++++++++++
End Sub

Private Sub Label1_Click()
lblpre = lblcl
lblcl = 1
Call lblclick
Label1.BackColor = RGB(0, 172, 193)
'heading label
Label10.Caption = Label1.Caption


End Sub

Private Sub Label11_Click()
lblpre = lblcl
lblcl = 11
Call lblclick
Label11.BackColor = RGB(0, 172, 193)
frmUser.Show
'heading label
Label10.Caption = Label11.Caption
End Sub

Private Sub Label2_Click()
lblpre = lblcl
lblcl = 2
Call lblclick
Label2.BackColor = RGB(0, 172, 193)
frmcustomer.Show
'heading label
Label10.Caption = Label2.Caption
End Sub

Private Sub Label3_Click()
lblpre = lblcl
lblcl = 3
Call lblclick
Label3.BackColor = RGB(0, 172, 193)
frmSupplier.Show
'heading label
Label10.Caption = Label3.Caption
End Sub

Private Sub Label4_Click()
lblpre = lblcl
lblcl = 4
Call lblclick
Label4.BackColor = RGB(0, 172, 193)
'heading label
Label10.Caption = Label4.Caption
frmSales.Show
End Sub

Private Sub Label5_Click()
lblpre = lblcl
lblcl = 5
Call lblclick
Label5.BackColor = RGB(0, 172, 193)
'frmPurchase.Show
Form2.Show

'heading label
Label10.Caption = Label5.Caption
End Sub

Private Sub Label6_Click()
lblpre = lblcl
lblcl = 6
Call lblclick
Label6.BackColor = RGB(0, 172, 193)
frmStocks.Show
'heading label
Label10.Caption = Label6.Caption
End Sub

Private Sub Label7_Click()
lblpre = lblcl
lblcl = 7
Call lblclick
Label7.BackColor = RGB(0, 172, 193)
'heading label
Label10.Caption = Label7.Caption
frmReport.Show
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
Case 11:
Label11.BackColor = RGB(51, 49, 44)
End Select

         

End Sub
