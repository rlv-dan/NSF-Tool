Attribute VB_Name = "Module"

'API to open dialog box with multiselect
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000                         '  new look commdlg
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHOWHELP = &H10

Public Type OPENFILENAME
        lStructSize As Long
        hWndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Public Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'''''''''''''''

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

'hand-pointer for links
Public Const IDC_HAND = 32649&
Public Const IDC_ARROW = 32512&
Public Declare Function LoadCursor Lib "USER32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "USER32" (ByVal hCursor As Long) As Long

Public nsfFormat
Public ConfirmRename As Boolean

Public Type NSF
    Path As String
    Title As String
    Artist As String
    Copyright As String
    Songs As Integer
    Specials As Integer
    System As Integer       'PAL/NTSC
    NSFE As Boolean
    Ripper As String
End Type

Public special_bit0 As String
Public special_bit1 As String
Public special_bit2 As String
Public special_bit3 As String
Public special_bit4 As String
Public special_bit5 As String

Public unknown_artist As String
Public unknown_title As String
Public unknown_copyright As String
Public unknown_ripper As String
Public ntsc_string As String
Public pal_string As String
Public ntscpal_string As String

Public sLastDir As String

Public Sub NsfRename()

    maxx = frmMain.ListView.ListItems.Count
    If maxx = 0 Then Exit Sub
    
    Dim myNSF As NSF
    Dim NoRemove As Boolean
    Dim curr As Integer
    curr = 1    'take first item
    
    For num = 1 To maxx
    
        NoRemove = False
        myNSF.Path = frmMain.ListView1.ListItems.Item(curr)
        ret = LoadNSF(myNSF)
        If ret = True Then
            
            If myNSF.Path = RemoveFile(myNSF.Path) & FormatName(myNSF) Then GoTo cont
            
            ret = Dir(RemoveFile(myNSF.Path) & FormatName(myNSF))
            
            If ret = FormatName(myNSF) Then    'Duplicate filenames
                ret = MsgBox("The file '" & RemovePath(myNSF.Path) & "' could not be renamed to '" & FormatName(myNSF) & "'." & Chr(10) & Chr(10) & "The filename already exists; skipping...", vbCritical, "Error!")
                NoRemove = True
                GoTo cont
            End If
            
            On Error GoTo err
            
            If ConfirmRename = True Then
                newName = FormatName(myNSF)
                newName = InputBox("Ok to rename, Cancel to skip.", "Confirm new filename", newName)
                If UCase(Right(newName, 4)) <> ".NSF" Then newName = newName + ".nsf"
                If newName <> ".nsf" Then
                    Name myNSF.Path As RemoveFile(myNSF.Path) & newName
                End If
            Else
                Name myNSF.Path As RemoveFile(myNSF.Path) & FormatName(myNSF)
            End If
            
            GoTo cont
err:
            ret = MsgBox("The file '" & RemovePath(myNSF.Path) & "' could not be renamed! Skipping...", vbCritical, "Error!")
            NoRemove = True
cont:
            On Error GoTo 0
        End If
        
        If NoRemove = False Then
            frmMain.ListView.ListItems.Remove (curr)
            frmMain.ListView1.ListItems.Remove (curr)
        Else
            curr = curr + 1 'to continue in list; jump over skipped ones
        End If
    
        frmMain.StatusBar.SimpleText = "Working..."
        DoEvents
    Next
    
    frmMain.StatusBar.SimpleText = "Ready"
    
End Sub

Public Function FormatName(ByRef myNSF As NSF) As String

    'Format
    FormatName = ""
    For ch = 1 To Len(nsfFormat)
        If Mid(nsfFormat, ch, 1) <> "%" Then
            FormatName = FormatName + Mid(nsfFormat, ch, 1)
        Else
            ch = ch + 1
            Select Case UCase(Mid(nsfFormat, ch, 1))
            
            Case "T":
                If myNSF.Title = "<?>" Or myNSF.Title = "?" Or LCase(myNSF.Title) = "na" Or LCase(myNSF.Title) = "n/a" Then
                    FormatName = FormatName & unknown_title    'Replace <?> & ?
                Else
                    FormatName = FormatName & myNSF.Title
                End If
                
            Case "A":
                If myNSF.Artist = "<?>" Or myNSF.Artist = "?" Or LCase(myNSF.Artist) = "na" Or LCase(myNSF.Artist) = "n/a" Then
                    FormatName = FormatName & unknown_artist    'Replace <?> & ?
                Else
                    FormatName = FormatName & myNSF.Artist
                End If
            
            Case "C":
                If myNSF.Copyright = "<?>" Or myNSF.Copyright = "?" Or LCase(myNSF.Copyright) = "na" Or LCase(myNSF.Copyright) = "n/a" Then
                    FormatName = FormatName & unknown_copyright    'Replace <?> & ?
                Else
                    FormatName = FormatName & myNSF.Copyright
                End If
            
            Case "R":
                If myNSF.Ripper = "<?>" Or myNSF.Ripper = "?" Or LCase(myNSF.Ripper) = "na" Or LCase(myNSF.Ripper) = "n/a" Then
                    FormatName = FormatName & unknown_ripper    'Replace <?> & ?
                Else
                    FormatName = FormatName & myNSF.Ripper
                End If
            
            Case "S":
                FormatName = FormatName & myNSF.Songs
            
            Case "F":
                FormatName = FormatName & enterpetSpecials(myNSF)
            
            Case "V":
                If (myNSF.System And 1) = 0 Then FormatName = FormatName & ntsc_string
                If (myNSF.System And 1) = 1 Then FormatName = FormatName & pal_string
                If (myNSF.System And 2) = 1 Then FormatName = FormatName & ntscpal_string
    
            Case Else
                FormatName = FormatName & "%" & Mid(nsfFormat, ch, 1)
            
            End Select
        End If
        
    Next
    
    If UCase(Right(FormatName, 4)) = ".NSF" Then FormatName = Mid(FormatName, 1, Len(FormatName) - 4)
        
    FormatName = Trim(FormatName)
    FormatName = Replace(FormatName, "  ", " ")
    FormatName = Replace(FormatName, Chr(34), "'")
    
    'Remove invalid characters: \ / : * ? < > " |
    Temp = FormatName
    FormatName = ""
    For num = 1 To Len(Temp)
        ch = Mid(Temp, num, 1)
        If ch <> "\" And ch <> "/" And ch <> ":" And ch <> "*" And ch <> "?" And ch <> "<" And ch <> ">" And ch <> "|" And ch <> Chr(34) Then
            FormatName = FormatName + ch
        End If
    Next
    
    'Add extension
    FormatName = FormatName + ".nsf"

End Function

Public Function LoadNSF(ByRef myNSF As NSF) As Boolean

    myPath = myNSF.Path
    
    myNSF.Title = ""
    myNSF.Copyright = ""
    myNSF.Artist = ""
    myNSF.Songs = 1
    myNSF.Specials = 0
    myNSF.NSFE = False
    myNSF.Ripper = ""
    
    ret = Dir(myPath)
    If ret = "" Then GoTo err
    
    'load nsf
    
    Open myPath For Binary As #1
    
        If Input(1, #1) <> "N" Then GoTo err
        If Input(1, #1) <> "E" Then GoTo err
        If Input(1, #1) <> "S" Then GoTo err
        If Input(1, #1) <> "M" Then GoTo err
        If Asc(Input(1, #1)) <> 26 Then GoTo err
    
        'version = Asc(Input(1, #1))
        Seek #1, (Seek(1) + 1)
        myNSF.Songs = Asc(Input(1, #1))
        Seek #1, (Seek(1) + 7)
        
        For a = 1 To 32
            myNSF.Title = myNSF.Title + Chr(Asc(Input(1, #1)))
        Next
        For a = 1 To 32
            myNSF.Artist = myNSF.Artist + Chr(Asc(Input(1, #1)))
        Next
        For a = 1 To 32
            myNSF.Copyright = myNSF.Copyright + Chr(Asc(Input(1, #1)))
        Next
    
        Seek #1, (Seek(1) + 12)
        myNSF.System = Asc(Input(1, #1))
    
        myNSF.Specials = Asc(Input(1, #1))
    
    
        Call TrimAndFix(myNSF)
    
    Close #1
    LoadNSF = True
    Exit Function
    
err:
    Close #1
    LoadNSF = False
    
    'try nsfe instead
    Open myPath For Binary As #1
    
        If Input(1, #1) <> "N" Then GoTo err2
        If Input(1, #1) <> "S" Then GoTo err2
        If Input(1, #1) <> "F" Then GoTo err2
        If Input(1, #1) <> "E" Then GoTo err2
        
        myNSF.NSFE = True
        
        nsfeData = Input(LOF(1) - Loc(1), #1)
        
        
        infoPos = InStr(1, nsfeData, "INFO")
        If infoPos > 4 Then
            'WORD   Load Address
            'WORD   Init Address
            'WORD   Play Address
            'BYTE   PAL/NTSC
            'BYTE   Extra sound chip support    'upp tom denna är mandatory (8 bytes)
            'BYTE   Number of Tracks
            'BYTE   Initial Track
    
            iChunkSizeString = Mid(nsfeData, infoPos - 4, 4)  'UINT = an unsigned integer.. 4 bytes in size, stored low byte first: 08 06 02 FF = 0xFF020608
            iChunkSize = Asc(Mid(iChunkSizeString, 1, 1)) + (Asc(Mid(iChunkSizeString, 2, 1)) * 256) + (Asc(Mid(iChunkSizeString, 3, 1)) * 65536) + (Asc(Mid(iChunkSizeString, 4, 1)) * 16777216)
    
            myNSF.System = Asc(Mid(nsfeData, infoPos + 4 + 6, 1))
    '           if bit 0 is 0 -> NSF is NTSC
    '           if bit 0 is 1   -> NSF is PAL
    '           if bit 1 is 1   -> Ignore bit 0, NSF is a Dual PAL/NTSC
    '           bits 2-7          -> Unknown.  Should be zero to allow for future
    '                                expansion.
    
            myNSF.Specials = Asc(Mid(nsfeData, infoPos + 4 + 7, 1)) 'same as nsf
    
            If iChunkSize > 8 Then
                myNSF.Songs = Asc(Mid(nsfeData, infoPos + 4 + 8, 1))
            End If
    
    
        Else
            GoTo err2
        End If
        
        
        authPos = InStr(1, nsfeData, "auth")    'not required
        If authPos > 4 Then
            iChunkSizeString = Mid(nsfeData, authPos - 4, 4)  'UINT = an unsigned integer.. 4 bytes in size, stored low byte first: 08 06 02 FF = 0xFF020608
            iChunkSize = Asc(Mid(iChunkSizeString, 1, 1)) + (Asc(Mid(iChunkSizeString, 2, 1)) * 256) + (Asc(Mid(iChunkSizeString, 3, 1)) * 65536) + (Asc(Mid(iChunkSizeString, 4, 1)) * 16777216)
            
            authData = Mid(nsfeData, authPos + 4, iChunkSize)
            authData = Split(authData, Chr(0))
            
            If UBound(authData) >= 1 Then myNSF.Title = authData(0)
            If UBound(authData) >= 2 Then myNSF.Artist = authData(1)
            If UBound(authData) >= 3 Then myNSF.Copyright = authData(2)
            If UBound(authData) >= 4 Then myNSF.Ripper = authData(3)    'nsfe only
    
            Call TrimAndFix(myNSF)
    
        End If
    
    Close #1
    
    LoadNSF = True
    Exit Function
    
    
err2:
    Close #1
    LoadNSF = False


End Function

Public Function RemoveFile(ByVal Path) As String

    If Right(Path, 1) = "\" Then Path = (Left(Path, Len(Path) - 1))
    
    For num = Len(Path) To 1 Step -1
        If Mid(Path, num, 1) = "\" Then Exit For
    Next
    
    RemoveFile = Left(Path, num)

End Function


Public Function RemovePath(ByVal Path) As String

    If Right(Path, 1) = "\" Then Path = (Left(Path, Len(Path) - 1))
    
    For num = Len(Path) To 1 Step -1
        If Mid(Path, num, 1) = "\" Then Exit For
    Next
    RemovePath = Right(Path, Len(Path) - num)

End Function

Public Sub AddFiles()

    ' api open dialog box '''''''''
    Dim tOPENFILENAME As OPENFILENAME
    Dim lResult As Long
    Dim vFiles As Variant
    Dim lIndex As Long, lStart As Long

    With tOPENFILENAME
        .flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES
        .hWndOwner = hWnd
        .nMaxFile = 8192    'max 32k??
        .lpstrFilter = "NSF Files (*.nsf)" & Chr(0) & "*.nsf" & Chr(0) & "All Files" & Chr(0) & "*.*" & Chr(0) & Chr(0)
        .lpstrFile = Space(.nMaxFile - 1) & Chr(0)
        .lpstrInitialDir = sLastDir
        .lStructSize = Len(tOPENFILENAME)
    End With

    lResult = GetOpenFileName(tOPENFILENAME)

    If lResult > 0 Then
        With tOPENFILENAME
            vFiles = Split(Left(.lpstrFile, InStr(.lpstrFile, Chr(0) & Chr(0)) - 1), Chr(0))
        End With

        frmMain.StatusBar.SimpleText = "Adding files..."
    
        Dim myNSF As NSF
        If UBound(vFiles) = 0 Then    'only one file
                myNSF.Path = vFiles(0)
                sLastDir = RemoveFile(vFiles(0))
                If CheckForDuplicateEntry(myNSF) = False Then   'Not in list
                    ret = LoadNSF(myNSF)
                    If ret = True Then
                        Set itmX = frmMain.ListView.ListItems.Add(, , , , 1)
                        itmX.Text = RemovePath(myNSF.Path)
                        itmX.SubItems(1) = FormatName(myNSF)
    
                        Set itmX2 = frmMain.ListView1.ListItems.Add()
                        itmX2.Text = myNSF.Path

                    End If
                End If
    
        Else    'multiple files
    
            myPath = vFiles(0)
            If Right$(myPath, 1) <> "\" Then myPath = myPath & "\"
            sLastDir = myPath
            
            myFile = ""
            For lIndex = 1 To UBound(vFiles)
            
                myNSF.Path = myPath & "\" & vFiles(lIndex)
                If CheckForDuplicateEntry(myNSF) = False Then   'Not in list
                    ret = LoadNSF(myNSF)
                    If ret = True Then
                        Set itmX = frmMain.ListView.ListItems.Add(, , , , 1)
                        itmX.Text = RemovePath(myNSF.Path)
                        itmX.SubItems(1) = FormatName(myNSF)

                        Set itmX2 = frmMain.ListView1.ListItems.Add()
                        itmX2.Text = myNSF.Path

                    End If
                End If
                
                DoEvents
            Next
        End If

    End If
    
    frmMain.StatusBar.SimpleText = "Showing " & frmMain.ListView.ListItems.Count & " files"

    If frmMain.ListView.ListItems.Count > 0 Then
        frmMain.Toolbar.Buttons.Item(2).Enabled = True
        frmMain.Toolbar.Buttons.Item(3).Enabled = True
        frmMain.Toolbar.Buttons.Item(5).Enabled = True
    End If

End Sub


Public Sub UpdateListview()

    Dim myNSF As NSF

    frmMain.ListView.ListItems.Clear
    
    For a = 1 To frmMain.ListView1.ListItems.Count
    
            myNSF.Path = frmMain.ListView1.ListItems.Item(a)
            ret = LoadNSF(myNSF)
            If ret = True Then
                Set itmX = frmMain.ListView.ListItems.Add(, , , , 1)
                itmX.Text = RemovePath(myNSF.Path)
                itmX.SubItems(1) = FormatName(myNSF)
                            
            End If
            
    Next
    
    frmMain.StatusBar.SimpleText = "Showing " & frmMain.ListView.ListItems.Count & " files"

End Sub

Public Function enterpetSpecials(myNSF As NSF) As String

    enterpetSpecials = ""
    
    If (myNSF.Specials And 1) = 1 Then enterpetSpecials = enterpetSpecials + special_bit0
    If (myNSF.Specials And 2) = 2 Then enterpetSpecials = enterpetSpecials + special_bit1
    If (myNSF.Specials And 4) = 4 Then enterpetSpecials = enterpetSpecials + special_bit2
    If (myNSF.Specials And 8) = 8 Then enterpetSpecials = enterpetSpecials + special_bit3
    If (myNSF.Specials And 16) = 16 Then enterpetSpecials = enterpetSpecials + special_bit4
    If (myNSF.Specials And 32) = 32 Then enterpetSpecials = enterpetSpecials + special_bit5
    
    If myNSF.Specials = 0 Then enterpetSpecials = spNormal
    
End Function

Public Function SaveNsfInfo(myNSF As NSF) As Boolean

    num = frmMain.ListView.SelectedItem.Index
    myNSF.Path = frmMain.ListView1.ListItems.Item(num)
    
    Dim myByte As Byte
    
    ret = Dir(myNSF.Path)
    If ret <> RemovePath(myNSF.Path) Then GoTo err
    
    On Error GoTo err
    
    Open myNSF.Path For Binary As #1
    
        If Input(1, #1) <> "N" Then GoTo err
        If Input(1, #1) <> "E" Then GoTo err
        If Input(1, #1) <> "S" Then GoTo err
        If Input(1, #1) <> "M" Then GoTo err
        If Asc(Input(1, #1)) <> 26 Then GoTo err
    
        Seek #1, 15
        For ch = 1 To Len(myNSF.Title)
            myByte = Asc(Mid(myNSF.Title, ch, 1))
            Put #1, , myByte
        Next
        For ch = 1 To (32 - Len(myNSF.Title))
            Put #1, , CByte(&O0)    'NULL
        Next
    
        For ch = 1 To Len(myNSF.Artist)
            myByte = Asc(Mid(myNSF.Artist, ch, 1))
            Put #1, , myByte
        Next
        For ch = 1 To (32 - Len(myNSF.Artist))
            Put #1, , CByte(&O0)    'NULL
        Next
    
        For ch = 1 To Len(myNSF.Copyright)
            myByte = Asc(Mid(myNSF.Copyright, ch, 1))
            Put #1, , myByte
        Next
        For ch = 1 To (32 - Len(myNSF.Copyright))
            Put #1, , CByte(&O0)    'NULL
        Next
    
    Close #1
    SaveNsfInfo = True
    Exit Function
    
err:
    Close #1
    SaveNsfInfo = False
    ret = MsgBox("Failed to save file...", vbCritical, "Error")
    On Error GoTo 0

End Function

Public Function CheckForDuplicateEntry(myNSF As NSF) As Boolean

    For num = 1 To frmMain.ListView1.ListItems.Count
        If frmMain.ListView1.ListItems.Item(num) = myNSF.Path Then
            CheckForDuplicateEntry = True
            Exit Function
        End If
    Next
    
    CheckForDuplicateEntry = False
    
End Function

Public Sub TrimAndFix(ByRef myNSF As NSF)

    For c = 0 To 31
        myNSF.Title = Replace(myNSF.Title, Chr(c), " ")
        myNSF.Artist = Replace(myNSF.Artist, Chr(c), " ")
        myNSF.Copyright = Replace(myNSF.Copyright, Chr(c), " ")
        myNSF.Ripper = Replace(myNSF.Ripper, Chr(c), " ")
    Next

    myNSF.Title = Replace(myNSF.Title, Chr(255), " ")
    myNSF.Artist = Replace(myNSF.Artist, Chr(255), " ")
    myNSF.Copyright = Replace(myNSF.Copyright, Chr(255), " ")
    myNSF.Ripper = Replace(myNSF.Ripper, Chr(255), " ")
            
    myNSF.Title = Trim(myNSF.Title)
    myNSF.Artist = Trim(myNSF.Artist)
    myNSF.Copyright = Trim(myNSF.Copyright)
    myNSF.Ripper = Trim(myNSF.Ripper)
    
End Sub


Public Function SaveNsfExtended(myNSF As NSF) As Boolean

    ret = Dir(myNSF.Path)
    If ret <> RemovePath(myNSF.Path) Then GoTo err
    
    On Error GoTo err
    
    If Dir(myNSF.Path & ".tmp") <> "" Then Kill myNSF.Path & ".tmp"

    Dim myFile As String

    Open myNSF.Path For Binary As #1
        myFile = Input(LOF(1), #1)
    Close #1


    Dim myFile1 As String
    Dim myFile2 As String
    Dim newAuth As String

    authPos = InStr(1, myFile, "auth")    'not required

    If authPos > 4 Then
    
        iChunkSizeString = Mid(myFile, authPos - 4, 4)  'UINT = an unsigned integer.. 4 bytes in size, stored low byte first: 08 06 02 FF = 0xFF020608
        iChunkSize = Asc(Mid(iChunkSizeString, 1, 1)) + (Asc(Mid(iChunkSizeString, 2, 1)) * 256) + (Asc(Mid(iChunkSizeString, 3, 1)) * 65536) + (Asc(Mid(iChunkSizeString, 4, 1)) * 16777216)

        myFile1 = Mid(myFile, 1, authPos - 5)
        myFile2 = Mid(myFile, authPos + 4 + iChunkSize)

        newData = ""
        
        tmp = myNSF.Title
        If tmp = "" Then tmp = "<?>"
        newData = newData & tmp & Chr(0)
        
        tmp = myNSF.Artist
        If tmp = "" Then tmp = "<?>"
        newData = newData & tmp & Chr(0)

        tmp = myNSF.Copyright
        If tmp = "" Then tmp = "<?>"
        newData = newData & tmp & Chr(0)

        If myNSF.Ripper <> "" Then
            newData = newData & myNSF.Ripper & Chr(0)
        End If
        
        newAuth = Chr(Len(newData)) & Chr(0) & Chr(0) & Chr(0) & "auth" & newData
        
        
    Else
        Close #1
        SaveNsfExtended = False
        Exit Function
    End If

    Open myNSF.Path & ".tmp" For Binary As #1
        Put #1, , myFile1 & newAuth & myFile2
    Close #1


    Kill myNSF.Path
    Name myNSF.Path & ".tmp" As myNSF.Path
    
    
    SaveNsfExtended = True
    Exit Function
    
err:
    Close #1
    SaveNsfExtended = False
    If Dir(myNSF.Path & ".tmp") <> "" Then Kill myNSF.Path & ".tmp"
    ret = MsgBox("Failed to save file...", vbCritical, "Error")
    On Error GoTo 0
    
End Function

