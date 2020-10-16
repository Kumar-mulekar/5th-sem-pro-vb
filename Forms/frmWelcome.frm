VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "WELCOME - - - 8-bit"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3480
      Top             =   5520
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   7575
      Left            =   -1800
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   -1  'True
      windowlessVideo =   -1  'True
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   21616
      _cy             =   13361
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public imgtm As Integer

Private Sub Form_Load()
Timer1.Enabled = True
wmp1.URL = App.Path + "\img_res\11.gif"


End Sub

Private Sub Image1_Click()

End Sub

Private Sub Timer1_Timer()
imgtm = imgtm + 1

If imgtm > 35 Then
  Form1.Show
  Form3.Hide
  Unload Me

End If


End Sub

