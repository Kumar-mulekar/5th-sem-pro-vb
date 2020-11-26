Attribute VB_Name = "Module1"
Public skn1 As SkinFramework
Public userName As String
Public Sub skin()



skn1.LoadSkin App.Path + "\Styles\Office2010.cjstyles", ""
skn1.ApplyWindow All.hWnd
End Sub

Public Sub ValidNumeric(KeyAscii As Integer)
'allow only numeric value
'Check whether the Input is numeric or not
Select Case KeyAscii
Case 8
Case 48 To 57
Case 47
Case 13
Case 32
Case 48 To 57
 Case Else
  MsgBox "Invalid Input.Please Enter Numeric Types Only..", vbOKOnly + vbExclamation
  KeyAscii = 0
End Select
End Sub
