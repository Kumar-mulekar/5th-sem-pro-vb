VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#18.6#0"; "CO23AE~1.OCX"
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12840
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   ScaleHeight     =   12840
   ScaleWidth      =   15255
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   1215
      Left            =   4800
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   16
      Top             =   7680
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   9480
      TabIndex        =   15
      Top             =   6120
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   5160
      TabIndex        =   14
      Top             =   5760
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   5520
      TabIndex        =   13
      Top             =   4680
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   8280
      TabIndex        =   12
      Top             =   4560
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   10800
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   1455
      Left            =   9360
      TabIndex        =   10
      Top             =   2640
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   855
      Left            =   6360
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   615
      Left            =   6480
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   9600
      TabIndex        =   7
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   6480
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "xp"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "2010"
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Codejock"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2007"
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   3000
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "vista"
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   960
      Picture         =   "Form1.frx":007A
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   5
      Top             =   480
      Width           =   2295
   End
   Begin XtremeSkinFramework.SkinFramework skn 
      Left            =   240
      Top             =   480
      _Version        =   1179654
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
skn.LoadSkin App.Path + "\Styles\Vista.cjstyles", ""
skn.ApplyWindow Me.hWnd
End Sub

Private Sub Command2_Click()
skn.LoadSkin App.Path + "\Styles\Office2007.cjstyles", ""
skn.ApplyWindow Me.hWnd
End Sub

Private Sub Command3_Click()
skn.LoadSkin App.Path + "\Styles\Codejock.cjstyles", ""
skn.ApplyWindow Me.hWnd
End Sub

Private Sub Command4_Click()
skn.LoadSkin App.Path + "\Styles\Office2010.cjstyles", ""
skn.ApplyWindow Me.hWnd



End Sub

Private Sub Command5_Click()
skn.LoadSkin App.Path + "\Styles\WinXP.Royale.cjstyles", ""
skn.ApplyWindow Me.hWnd
End Sub

