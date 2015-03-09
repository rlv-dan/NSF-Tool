Attribute VB_Name = "IniReadWrite"
'Example:
'    Text1.Text = ReadIniValue(App.Path & "\MyTest.ini", "Default", "Text1", "Default Value")
'    WriteIniValue App.Path & "\MyTest.ini", "Default", "Text1", Text1.Text


Option Explicit

Public Function ReadIniValue(INIpath As String, KEY As String, Variable As String, DefaultValue As String) As String
    Dim NF As Integer
    Dim Temp As String
    Dim LcaseTemp As String
    Dim ReadyToRead As Boolean
        
AssignVariables:
            NF = FreeFile
            ReadIniValue = DefaultValue
            KEY = "[" & LCase$(KEY) & "]"
            Variable = LCase$(Variable)
        
EnsureFileExists:
        Open INIpath For Binary As NF
        Close NF
        SetAttr INIpath, vbArchive
        
LoadFile:
        Open INIpath For Input As NF
        While Not EOF(NF)
        Line Input #NF, Temp
        LcaseTemp = LCase$(Temp)
        If InStr(LcaseTemp, "[") <> 0 Then ReadyToRead = False
        If LcaseTemp = KEY Then ReadyToRead = True
        If InStr(LcaseTemp, "[") = 0 And ReadyToRead = True Then
            If InStr(LcaseTemp, Variable & "=") = 1 Then
                ReadIniValue = Mid$(Temp, 1 + Len(Variable & "="))
                Close NF: Exit Function
                End If
            End If
        Wend
        Close NF
End Function

Public Function WriteIniValue(INIpath As String, PutKey As String, PutVariable As String, PutValue)
    Dim Temp As String
    Dim LcaseTemp As String
    Dim ReadKey As String
    Dim ReadVariable As String
    Dim LOKEY As Integer
    Dim HIKEY As Integer
    Dim KEYLEN As Integer
    Dim VAR As Integer
    Dim VARENDOFLINE As Integer
    Dim NF As Integer
    Dim X As Integer

    Dim sPutValue As String
    sPutValue = CStr(PutValue)

AssignVariables:
        NF = FreeFile
        ReadKey = vbCrLf & "[" & LCase$(PutKey) & "]" & Chr$(13)
        KEYLEN = Len(ReadKey)
        ReadVariable = Chr$(10) & LCase$(PutVariable) & "="
            
EnsureFileExists:
        Open INIpath For Binary As NF
        Close NF
        SetAttr INIpath, vbArchive
        
LoadFile:
        Open INIpath For Input As NF
        Temp = Input$(LOF(NF), NF)
        Temp = vbCrLf & Temp & "[]"
        Close NF
        LcaseTemp = LCase$(Temp)
        
LogicMenu:
        LOKEY = InStr(LcaseTemp, ReadKey)
        If LOKEY = 0 Then GoTo AddKey:
        HIKEY = InStr(LOKEY + KEYLEN, LcaseTemp, "[")
        VAR = InStr(LOKEY, LcaseTemp, ReadVariable)
        If VAR > HIKEY Or VAR < LOKEY Then GoTo AddVariable:
        GoTo RenewVariable:
        
AddKey:
            Temp = Left$(Temp, Len(Temp) - 2)
            Temp = Temp & vbCrLf & vbCrLf & "[" & PutKey & "]" & vbCrLf & PutVariable & "=" & sPutValue
            GoTo TrimFinalString:
            
AddVariable:
            Temp = Left$(Temp, Len(Temp) - 2)
            Temp = Left$(Temp, LOKEY + KEYLEN) & PutVariable & "=" & sPutValue & vbCrLf & Mid$(Temp, LOKEY + KEYLEN + 1)
            GoTo TrimFinalString:
            
RenewVariable:
            Temp = Left$(Temp, Len(Temp) - 2)
            VARENDOFLINE = InStr(VAR, Temp, Chr$(13))
            Temp = Left$(Temp, VAR) & PutVariable & "=" & sPutValue & Mid$(Temp, VARENDOFLINE)
            GoTo TrimFinalString:
    
TrimFinalString:
            Temp = Mid$(Temp, 2)
            Do Until InStr(Temp, vbCrLf & vbCrLf & vbCrLf) = 0
            Temp = Replace(Temp, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf)
            Loop
        
            Do Until Right$(Temp, 1) > Chr$(13)
            Temp = Left$(Temp, Len(Temp) - 1)
            Loop
        
            Do Until Left$(Temp, 1) > Chr$(13)
            Temp = Mid$(Temp, 2)
            Loop
        
OutputAmendedINIFile:
            Open INIpath For Output As NF
            Print #NF, Temp
            Close NF
        
End Function

