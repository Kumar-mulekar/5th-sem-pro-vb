VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'=========================================================================================
'Allow only alfabets in text box
'1'st Argument, KeyAscii - It will indicate the Key Pressed
'                         by User
'2'nd Argument, mbBlankAllowFlag - It is for Checking that
'                                 the Blank Space in text box
'                                 is allowed or not
' Ascii 32 is for Blank Space
' Ascii 27 is for Esc
Public Function OnlyAlfabets(KeyAscii As Integer, mbBlankAllowFlag As Boolean, Optional blnMSGBOX As Boolean) As Integer
    If (KeyAscii = vbKeyBack Or KeyAscii = Asc(".") Or (KeyAscii = 32 And mbBlankAllowFlag) Or KeyAscii = vbKeyReturn Or KeyAscii = 27) Then
        OnlyAlfabets = KeyAscii
    ElseIf ((KeyAscii < vbKeyA Or KeyAscii > vbKeyZ) And (KeyAscii < Asc("a") Or KeyAscii > Asc("z"))) Then
            Beep
            If blnMSGBOX = True Then MsgBox "Only Alphabets are allowed.", vbInformation, "Invalid Alphabet"
            OnlyAlfabets = 0
        Else
            OnlyAlfabets = KeyAscii
    End If
End Function
'=========================================================================================
'Allow only numeric keys in text fields
'1'st Argument, KeyAscii - It will indicate the Key Pressed
'                         by User
Public Function OnlyNumeric(KeyAscii As Integer, mbFractionAllowFlag As Boolean, Optional bSupressMsg As Boolean) As Integer
    If (KeyAscii = vbKeyBack Or (KeyAscii = Asc(".") And mbFractionAllowFlag) Or KeyAscii = vbKeyReturn Or KeyAscii = 27) Then
        OnlyNumeric = KeyAscii
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
        Beep
        If bSupressMsg = True Then MsgBox "Only Numbers are allowed", vbInformation, "Invalid Number"
        OnlyNumeric = 0
        Else
        OnlyNumeric = KeyAscii
    End If
End Function
'=========================================================================================
'Allow both alpabets and numbers reject all other
'1'st Argument, KeyAscii - It will indicate the Key Pressed
'                         by User

'2'nd Argument, mbBlankAllowFlag - It is for Checking that
'                                 the Blank Space in text box
'                                 is allowed or not

Public Function OnlyAlfaNumeric(KeyAscii As Integer, mbBlankAllowFlag As Boolean, Optional blnMSGBOX As Boolean) As Integer
    If (KeyAscii = vbKeyBack Or KeyAscii = Asc(".") Or (KeyAscii = 32 And mbBlankAllowFlag) Or KeyAscii = vbKeyReturn Or KeyAscii = 27) Then
        OnlyAlfaNumeric = KeyAscii
    ElseIf ((KeyAscii < vbKeyA Or KeyAscii > vbKeyZ) _
            And (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) _
            And (KeyAscii < vbKey0 Or KeyAscii > vbKey9)) Then
            Beep
            If blnMSGBOX Then MsgBox "Only Alphabets and Numbers are allowed", vbInformation, "Invalid Input"
            OnlyAlfaNumeric = 0
        Else
            OnlyAlfaNumeric = KeyAscii
    End If
End Function
'=========================================================================================
'Allow only date related characters in text fields
'1'st Argument, KeyAscii - It will indicate the Key Pressed
'                         by User
Public Function OnlyDate(KeyAscii As Integer, Optional blnMSGBOX As Boolean) As Integer
    If (KeyAscii = vbKeyBack Or KeyAscii = Asc("/") Or KeyAscii = vbKeyReturn Or KeyAscii = 27) Then 'Or KeyAscii = Asc("-")
        OnlyDate = KeyAscii
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
            Beep
            If blnMSGBOX Then MsgBox "Only date related Characters (DD/MM/YYYY) are allowed.", vbInformation, "Invalid Date"
            OnlyDate = 0
        Else
            OnlyDate = KeyAscii
    End If
End Function
'=========================================================================================
'Allow only Time related characters in text fields
'1'st Argument, KeyAscii - It will indicate the Key Pressed
'                         by User
Public Function OnlyTime(KeyAscii As Integer, Optional blnMSGBOX As Boolean) As Integer
    If (KeyAscii = vbKeyBack Or KeyAscii = Asc(".") Or KeyAscii = vbKeyReturn Or KeyAscii = 27) Then
        OnlyTime = KeyAscii
    ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
            Beep
            If blnMSGBOX Then MsgBox "Only Time related Characters (HH.MM) is allowed.", vbInformation, "Invalid Time"
            OnlyTime = 0
        Else
            OnlyTime = KeyAscii
    End If
End Function

'=========================================================================================
'Allow only Phone No. related characters in text fields
'1'st Argument, KeyAscii - It will indicate the Key Pressed
'                         by User
'In Phone No. the Numbers 0-9 & Character of type -,/, (, ) are only allowed.
' Ascii 32 is for BlankSpace, 27 for Esc
Public Function OnlyPhoneNo(KeyAscii As Integer, Optional blnMSGBOX As Boolean) As Integer
    If (KeyAscii = vbKeyBack Or KeyAscii = 32 Or KeyAscii = vbKeyReturn _
                             Or KeyAscii = 27) Or KeyAscii = Asc(",") _
                             Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") _
                             Or KeyAscii = Asc(")") Or KeyAscii = Asc("/") Then
        OnlyPhoneNo = KeyAscii
        ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
            Beep
            If blnMSGBOX Then MsgBox "Only Phone No. Related Charactrs are allowed.", vbInformation, "Invalid Phone No."
            OnlyPhoneNo = 0
        Else
            OnlyPhoneNo = KeyAscii
    End If
End Function
'=========================================================================================
'Allow only E-Mail. related characters in text fields
'1'st Argument, KeyAscii - It will indicate the Key Pressed
'                         by User
'In E-Mail the Numbers 0-9 & Character of type A-Z,a-z,_,.,@,/ are only allowed.
' Ascii 32 is for BlankSpace, 27 for Esc
Public Function OnlyEMail(KeyAscii As Integer, Optional blnMSGBOX As Boolean) As Integer
    
    If (KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Or KeyAscii = 27 _
        Or KeyAscii = Asc("-") Or KeyAscii = Asc("_") Or KeyAscii = Asc(".") _
        Or KeyAscii = Asc("@") Or KeyAscii = Asc("/")) Then
        OnlyEMail = KeyAscii
    ElseIf (KeyAscii < vbKeyA Or KeyAscii > vbKeyZ) And _
           (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) And _
           (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
            Beep
            If blnMSGBOX Then MsgBox "Only E-Mail related characters are allowed.", vbInformation, "Invalid E-Mail"
            OnlyEMail = 0
        Else
            OnlyEMail = KeyAscii
    End If
End Function

'=========================================================================================
' Check the Date
'1'st Argument, KeyAscii - It will indicate the Key Pressed
'                         by User
' Check the Date given by user is correct or not & Convert it in to British format
' Ascii 32 is for BlankSpace, 27 for Esc
Public Function CheckDate(txtBox As ComboBox) As Boolean
    Dim Dobdate$
    If (Len(txtBox) = 10 And IsNumeric(Mid(txtBox, 4, 2))) Then
        If (Val(Mid(txtBox, 4, 2)) <= 12 And Val(Mid(txtBox, 4, 2)) > 0 And Val(Mid(txtBox, 1, 2)) > 0 And Val(Mid(txtBox, 1, 2)) < 32 And (IsDate(txtBox))) Then
            'Convert the British formatted date to American for validation
            'of system date greater than birthdate
            Dobdate$ = Mid(txtBox, 4, 2) & "/" & Mid(txtBox, 1, 2) & "/" & Mid(txtBox, 7, 4)
             
            If Val(Mid(txtBox, 7, 4)) >= 1900 Then
                CheckDate = True
                Exit Function
            Else
                MsgBox "The date is Less 01/01/1900. It should be greater or equal to than 01/01/1900", vbInformation, "Date Error"
            End If
        Else
            MsgBox "Invalid Date format.", vbInformation, "Invalid Date"
        End If
    Else
        
        MsgBox "Date must be in DD/MM/YYYY format .", vbInformation, "Invalid Date"
    End If
    

End Function

'=========================================================================================
' Check the Time
'1'st Argument, KeyAscii - It will indicate the Key Pressed
'                         by User
' Check the Time given by user
' Ascii 32 is for BlankSpace, 27 for Esc
Public Function CheckTime(txtBox As TextBox) As Boolean
    Dim DobTime$, DobHrs$, DobMin$, DobSec$
    txtBox.Text = Trim(txtBox.Text)
    If IsDate(txtBox.Text) Then
        txtBox.Text = Format(txtBox.Text, "HH:MM:SS")
        CheckTime = True
    Else
        MsgBox "Time must be in HH:MM:SS format.", vbInformation, "Invalid Time"
    End If
End Function

Public Function NewCheckDate(txtBox As ComboBox) As Boolean
    Dim Dobdate$
    If (Len(txtBox) = 10 And IsNumeric(Mid(txtBox, 4, 2))) Then
        If (Val(Mid(txtBox, 4, 2)) <= 12 And Val(Mid(txtBox, 4, 2)) > 0 And Val(Mid(txtBox, 1, 2)) > 0 And Val(Mid(txtBox, 1, 2)) < 32 And (IsDate(txtBox))) Then
            'Convert the British formatted date to American for validation
            'of system date greater than birthdate
            Dobdate$ = Mid(txtBox, 4, 2) & "/" & Mid(txtBox, 1, 2) & "/" & Mid(txtBox, 7, 4)
            
            If Date >= CDate(Dobdate$) And Val(Mid(txtBox, 7, 4)) >= 1900 Then
                NewCheckDate = True
                Exit Function
            Else
                MsgBox "This date exceeds present Date or Its year is Less 1900.", vbInformation, "Date Error"
            End If
        Else
            MsgBox "Invalid Date format.", vbInformation, "Invalid Date"
        End If
    Else
        MsgBox "Date must be in DD/MM/YYYY format.", vbInformation, "Invalid Date"
    End If
End Function

Public Function readResultFile(sfileName$) As Boolean
    Dim filebuffer As String
    Dim txtFile As String
    Dim dFileHdl As Double
    
     'dFileHdl = Shell(App.Path & "\db\" & sfileName & " " & gComPort, vbMaximizedFocus)
     'MsgBox App.Path & "\db\" & sfileName & " " & gComPort
     'dFileHdl = Shell("D:\VBKSRTC\DB\Upld.exe 1")
    '**/
    On Error GoTo ErrReadFile
    txtFile = App.Path & "\db\result.txt"
    'Open "result.txt" For Input As #1     ' Open file for input parameters
    Open txtFile For Input As #1
    Do While Not EOF(1)
        Input #1, filebuffer
        If (Left(filebuffer, 2) = "OK") Then
            readResultFile = True
        Else
            MsgBox filebuffer, vbCritical, "Error authorising machine"
            readResultFile = False
        End If
    Loop
    Close #1
    Exit Function
ErrReadFile:
    MsgBox "ReturnResult  .txt file not created" & vbCrLf
    readResultFile = False
End Function
'=========================================================================================
Public Function FormatDate(dt$, rotate As Boolean) As String
    Dim iSlash1%, iSlash2%, Dobdate$
    If IsDate(dt$) Then
        Dim sDD$, sMM$, sYYYY$
        iSlash1 = InStr(1, dt$, "/")
        sDD = Left(dt$, iSlash1 - 1)
        iSlash2 = InStr(iSlash1 + 1, dt$, "/")
        sMM = Mid(dt$, iSlash1 + 1, (iSlash2 - iSlash1) - 1)
        sYYYY = Mid(dt$, iSlash2 + 1, Len(dt$))
        If (rotate) Then
            FormatDate = sDD$ & "/" & sMM$ & "/" & sYYYY$
        Else
            FormatDate = sMM$ & "/" & sDD$ & "/" & sYYYY$
        End If
    Else
        MsgBox "Invaid Date Format.", vbInformation, "Invalid Date"
    End If
End Function
Public Function FormatTime(txtBox As TextBox) As Boolean
    Dim sHH$, sMM$, iSemiColon1, iSemiColon2
    txtBox.Text = Trim(txtBox.Text)
    If Not IsNull(txtBox.Text) Then
        If Mid(txtBox, 2, 1) = "." Or Mid(txtBox, 2, 1) = ":" Then
            sHH$ = Mid(txtBox.Text, 1, 1)
            sMM$ = Mid(txtBox.Text, 3, 2)
        Else
            sHH$ = Mid(txtBox.Text, 1, 2)
            sMM$ = Mid(txtBox.Text, 4, 2)
        End If
        sHH$ = Format(sHH$, "00")
        sMM$ = Trim(sMM$)
        If Len(sMM$) = 0 Then
            sMM$ = "00"
        ElseIf Len(sMM$) = 1 Then
            sMM$ = sMM$ & "0"
        End If
        txtBox.Text = sHH$ & ":" & sMM$
        If IsDate(txtBox.Text) Then
            If sHH$ > 23 Or sMM$ > 59 Then
                MsgBox "In Time, Hours must be less than 24" & vbCrLf & " Min must be less than 60.", vbInformation, "Invalid Time"
                FormatTime = False
            Else
                FormatTime = True
            End If
        End If
    Else
        MsgBox "Time must be in HH.MM format.", vbInformation, "Invalid Time"
        FormatTime = False
    End If
End Function


' Function for making the initial Alphabets of words in Upper Case &_
' Other alphabets in lower case, it will take one argument that is
' to be Capitalized & make it proper. It will return that properly
' formatted string.
Public Function MakeInitialUpper(str As String) As String
        
        'Declarations: for Removing duplicate spaces within strings.
        Const TWO_SPACES As String = "  "
        Dim intPos As Integer
        Dim TempStr As String
        
        'Task: Convert all wordz with the first letter capitalized
        str = StrConv(str, vbProperCase)
        
        'Task: Removes duplicate spaces within strings.
        intPos = InStr(1, str, TWO_SPACES, vbBinaryCompare)
        Do While intPos > 0
            TempStr = LTrim$(Mid$(str, intPos + 1))
            str = Left$(str, intPos) & TempStr
            intPos = InStr(1, str, TWO_SPACES, vbBinaryCompare)
        Loop
        
        ' Return: The formated string
        MakeInitialUpper = str
End Function
' Function for Formatting the Passed String With Given parameter
' Parametr Description
'       1.   Str = String to be Formatted
'       2.  AlignFlag = Right or Left Align Flag
'       3.  CharFormat = Character that is applied for Format
'       4.  CharCount = No of Char Applied for Format
Public Function MyFormat(str As String, AlignFlag As String, CharFormat As String, CharCount As Byte, Optional ApplyDecimal As Boolean) As String
    Dim FormattedStr As String
    Dim i As Byte
    If ApplyDecimal Then
        i = InStr(1, str, ".", vbTextCompare)
        If i = 0 Then
            str = str & ".00"
        Else
            str = Format(str, "0.00")
        End If
        If CharCount - Len(str) + 3 > 0 Then
            For i = 1 To CharCount - Len(str) + 3
                    FormattedStr = FormattedStr & CharFormat
            Next
        End If
    Else
        If CharCount - Len(str) > 0 Then
            For i = 1 To CharCount - Len(str)
                    FormattedStr = FormattedStr & CharFormat
            Next
        End If
    End If
    If UCase(AlignFlag) = "R" Then
            FormattedStr = str & FormattedStr
    ElseIf UCase(AlignFlag) = "L" Then
        FormattedStr = FormattedStr & str
    End If
    MyFormat = FormattedStr
End Function
        
'=========================================================================================
'Allow only mobile No. related characters in text fields
'1'st Argument, KeyAscii - It will indicate the Key Pressed
'                         by User
'In Phone No. the Numbers 0-9 & Character of type -,/, (, ) are only allowed.
' Ascii 32 is for BlankSpace, 27 for Esc
Public Function OnlyMobileNo(KeyAscii As Integer, Optional blnMSGBOX As Boolean) As Integer
    If (KeyAscii = vbKeyBack Or KeyAscii = 32 Or KeyAscii = vbKeyReturn _
                             Or KeyAscii = 27) Or KeyAscii = Asc(",") _
                             Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") _
                             Or KeyAscii = Asc(")") Or KeyAscii = Asc("/") Then
        OnlyMobileNo = KeyAscii
        ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
            Beep
            If blnMSGBOX Then MsgBox "Only Mobile No. Related Charactrs are allowed.", vbInformation, "Invalid Mobile No."
            OnlyMobileNo = 0
        Else
            OnlyMobileNo = KeyAscii
    End If
   
End Function


'=========================================================================================
'Allow only mobile No. related characters in text fields
'1'st Argument, KeyAscii - It will indicate the Key Pressed
'                         by User
'In Phone No. the Numbers 0-9 & Character of type -,/, (, ) are only allowed.
' Ascii 32 is for BlankSpace, 27 for Esc
Public Function OnlyFaxNo(KeyAscii As Integer, Optional blnMSGBOX As Boolean) As Integer
    If (KeyAscii = vbKeyBack Or KeyAscii = 32 Or KeyAscii = vbKeyReturn _
                             Or KeyAscii = 27) Or KeyAscii = Asc(",") _
                             Or KeyAscii = Asc("-") Or KeyAscii = Asc("(") _
                             Or KeyAscii = Asc(")") Or KeyAscii = Asc("/") Then
        OnlyFaxNo = KeyAscii
        ElseIf (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
            Beep
            If blnMSGBOX Then MsgBox "Only Fax No. Related Charactrs are allowed.", vbInformation, "Invalid Fax No."
            OnlyFaxNo = 0
        Else
            OnlyFaxNo = KeyAscii
    End If
   
End Function


