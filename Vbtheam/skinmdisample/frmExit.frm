VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExit 
   BorderStyle     =   0  'None
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   LinkTopic       =   "Form4"
   ScaleHeight     =   2535
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   512
      ImageHeight     =   512
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExit.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExit.frx":187F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   2505
      Left            =   5640
      Picture         =   "frmExit.frx":27F5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2505
   End
   Begin VB.Image Image2 
      Height          =   2505
      Left            =   2880
      Picture         =   "frmExit.frx":D831
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2505
   End
   Begin VB.Image Image1 
      Height          =   2505
      Left            =   0
      Picture         =   "frmExit.frx":F0A0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2505
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Image1.Picture = "C:\Users\HYDRA\Desktop\cancel.gif"

End Sub

Private Sub Image2_Click()
End

End Sub
