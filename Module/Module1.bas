Attribute VB_Name = "Module1"
Public skn1 As SkinFramework
Public userName, userAccess As String
Public check As New InValidation
Public Sub skin()



skn1.LoadSkin App.Path + "\Styles\Office2010.cjstyles", ""
skn1.ApplyWindow All.hWnd
End Sub



Public Sub validateAN(KeyAscii As Integer)
    Dim i As Integer
    i = check.OnlyAlfaNumeric(KeyAscii, True, True)
    If i = 0 Then
        KeyAscii = 0
    End If
End Sub
Public Sub validateA(KeyAscii As Integer)
    Dim i As Integer
    i = check.OnlyAlfabets(KeyAscii, True, True)
    If i = 0 Then
        KeyAscii = 0
    End If
End Sub
Public Sub validateN(KeyAscii As Integer)
    Dim i As Integer
    i = check.OnlyNumeric(KeyAscii, True, True)
    If i = 0 Then
        KeyAscii = 0
    End If
End Sub
Public Sub validateE(KeyAscii As Integer)
    Dim i As Integer
    i = check.OnlyEMail(KeyAscii, True)
    If i = 0 Then
        KeyAscii = 0
    End If
End Sub

