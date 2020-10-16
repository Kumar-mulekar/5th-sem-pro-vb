Attribute VB_Name = "Module1"
Public skn1 As SkinFramework
Public Sub skin()



skn1.LoadSkin App.Path + "\Styles\Office2010.cjstyles", ""
skn1.ApplyWindow All.hWnd
End Sub

