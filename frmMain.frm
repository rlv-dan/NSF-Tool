VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "NSF Tool"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   5790
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3735
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
      Visible         =   0   'False
      _ExtentX        =   2566
      _ExtentY        =   873
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Path"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2520
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   63488
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3366
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3948
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":450C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4AEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   953
      ButtonWidth     =   1482
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Object.ToolTipText     =   "Add files"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Remove"
            Object.ToolTipText     =   "Remove selected"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Properties"
            Object.ToolTipText     =   "View information"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rename"
            Object.ToolTipText     =   "Rename"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Settings"
            Object.ToolTipText     =   "Settings"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Object.ToolTipText     =   "About"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Drop your NSF files here"
      Top             =   960
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Current Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Preview"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    Toolbar.ImageList = ImageList2
    Toolbar.Buttons.Item(1).Image = 1
    Toolbar.Buttons.Item(2).Image = 2
    Toolbar.Buttons.Item(3).Image = 3
    Toolbar.Buttons.Item(5).Image = 4
    Toolbar.Buttons.Item(7).Image = 5
    Toolbar.Buttons.Item(8).Image = 6
    
    Toolbar.Buttons.Item(2).Enabled = False
    Toolbar.Buttons.Item(3).Enabled = False
    Toolbar.Buttons.Item(5).Enabled = False
    
    ListView.TOp = Toolbar.TOp + Toolbar.Height + 80
    
    ''' settings '''
    nsfFormat = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "Format", "%T %F.nsf")
    ConfirmRename = CBool(ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "ConfirmRename", False))
    
    frmMain.Left = Val(ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "Left", "500"))
    frmMain.TOp = Val(ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "Top", "500"))
    frmMain.Width = Val(ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "Width", "5000"))
    frmMain.Height = Val(ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "Height", "4000"))
    col = Val(ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "Columns", "-1"))
    frmMain.WindowState = Val(ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "WindowState", "0"))
    
    sLastDir = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "LastDir", "c:\")
    
    special_bit0 = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "VRCVI", "- (Konami VRCVI)")
    special_bit1 = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "VRCVII", "- (Konami VRCVII)")
    special_bit2 = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "FDS", "- (FDS Sound)")
    special_bit3 = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "MMC5", "- (Nintendo MMC5 Audio)")
    special_bit4 = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "NAMCO", "- (Namco-106)")
    special_bit5 = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "SUNSOFT", "- (Sunsoft FME-07)")
    unknown_artist = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "UNKNOWN_ARTIST", "Unknown Artist")
    unknown_title = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "UNKNOWN_TITLE", "Unknown Title")
    unknown_copyright = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "UNKNOWN COPYRIGHT", "Unknown Copyright")
    unknown_ripper = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "UNKNOWN RIPPER", "Unknown Ripper")
    ntsc_string = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "NTSC", "NTSC")
    pal_string = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "PAL", "PAL")
    ntscpal_string = ReadIniValue(App.Path & "\Settings.ini", "NSF Tool", "DUAL PAL-NTSC", "Dual PAL-NTSC")
    
    ''''''''''''''''
    
    If col = -1 Then    'First time
        ListView.ColumnHeaders.Item(1).Width = (ListView.Width / 3) - 90
        ListView.ColumnHeaders.Item(2).Width = ListView.Width / 3 * 2
    Else
        ListView.ColumnHeaders.Item(1).Width = col
        ListView.ColumnHeaders.Item(2).Width = ListView.Width - col
    End If
    
    StatusBar.SimpleText = "Ready"
    
    'XP Themes
    If IsThemed() Then
        FixThemeSupport Controls
    End If

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    If frmMain.Width < 5910 Then frmMain.Width = 5910
    If frmMain.Height < 4620 Then frmMain.Height = 4620
    
    ListView.Height = frmMain.Height - 1500
    ListView.Width = frmMain.Width - 120
    
    ListView.ColumnHeaders.Item(2).Width = ListView.Width - ListView.ColumnHeaders.Item(1).Width - 50

End Sub


Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "Format", nsfFormat
    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "ConfirmRename", ConfirmRename
    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "LastDir", sLastDir
    
    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "Columns", ListView.ColumnHeaders.Item(1).Width
    
    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "VRCVI", special_bit0
    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "VRCVII", special_bit1
    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "FDS", special_bit2
    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "MMC5", special_bit3
    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "NAMCO", special_bit4
    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "SUNSOFT", special_bit5
    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "UNKNOWN_ARTIST", unknown_artist
    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "UNKNOWN_TITLE", unknown_title
    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "UNKNOWN COPYRIGHT", unknown_copyright
    
    Temp = frmMain.WindowState
    If Temp = 0 Then    'Only if normal state
        WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "Left", frmMain.Left
        WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "Top", frmMain.TOp
        WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "Width", frmMain.Width
        WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "Height", frmMain.Height
    End If
    If Temp = 1 Then Temp = 0   'Do not save minimized state
    WriteIniValue App.Path & "\Settings.ini", "NSF Tool", "WindowState", Temp
    
    Call UnloadXpApp

End Sub


Private Sub ListView_DblClick()

    If ListView.ListItems.Count > 0 Then frmInfo.Show vbModal

End Sub

Private Sub ListView_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case vbKeyInsert
        AddFiles
    
    Case vbKeyDelete
        For num = ListView.ListItems.Count To 1 Step -1
            If ListView.ListItems.Item(num).Selected = True Then
                ListView.ListItems.Remove (num)
                ListView1.ListItems.Remove (num)
            End If
        Next
        StatusBar.SimpleText = "Showing " & ListView.ListItems.Count & " file(s)"
    
        If frmMain.ListView.ListItems.Count = 0 Then
            Toolbar.Buttons.Item(2).Enabled = False
            Toolbar.Buttons.Item(3).Enabled = False
            Toolbar.Buttons.Item(5).Enabled = False
        End If
    
    End Select

End Sub

Private Sub ListView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)

    Dim myNSF As NSF
    
    If Data.GetFormat(15) = True Then
        num = 1
        frmMain.StatusBar.SimpleText = "Working..."
        frmMain.Enabled = False
        On Error GoTo err
        Do
            myNSF.Path = Data.Files(num)
            If CheckForDuplicateEntry(myNSF) = False Then   'Not in list
                ret = LoadNSF(myNSF)
                If ret = True Then
                    Set itmX = ListView.ListItems.Add(, , , , 1)
                    itmX.Text = RemovePath(myNSF.Path)
                    itmX.SubItems(1) = FormatName(myNSF)
                    
                    Set itmX2 = frmMain.ListView1.ListItems.Add()
                    itmX2.Text = myNSF.Path
        
                End If
            End If
            num = num + 1
            DoEvents
        Loop
        
err:
    
        On Error GoTo 0
        StatusBar.SimpleText = "Showing " & ListView.ListItems.Count & " file(s)"
        
            If frmMain.ListView.ListItems.Count > 0 Then
                Toolbar.Buttons.Item(2).Enabled = True
                Toolbar.Buttons.Item(3).Enabled = True
                Toolbar.Buttons.Item(5).Enabled = True
            End If
        
        frmMain.Enabled = True
    
    End If

End Sub

Private Sub mnuAdd_Click()

    CommonDialog.flags = cdlOFNAllowMultiselect + cdlOFNFileMustExist + cdlOFNPathMustExist + cdlOFNLongNames + cdlOFNExplorer
    CommonDialog.Filter = "NSF Files (*.nsf)|*.nsf|All Files|*.*"
    CommonDialog.ShowOpen

End Sub

Private Sub mnuSettings_Click()

    frmSettings.Show vbModal

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    Case 1:
        AddFiles
        
    Case 2:
        If ListView.ListItems.Count > 0 Then
            For num = ListView.ListItems.Count To 1 Step -1
                If ListView.ListItems.Item(num).Selected = True Then
                    ListView.ListItems.Remove (num)
                    ListView1.ListItems.Remove (num)
                End If
            Next
            StatusBar.SimpleText = "Showing " & ListView.ListItems.Count & " file(s)"
            If frmMain.ListView.ListItems.Count = 0 Then
                Toolbar.Buttons.Item(2).Enabled = False
                Toolbar.Buttons.Item(3).Enabled = False
                Toolbar.Buttons.Item(5).Enabled = False
            End If
        End If
        
    Case 3:
        If ListView.ListItems.Count > 0 Then frmInfo.Show vbModal
        
    Case 5:
        Me.Enabled = False
        NsfRename
        Me.Enabled = True
        If frmMain.ListView.ListItems.Count = 0 Then
            Toolbar.Buttons.Item(2).Enabled = False
            Toolbar.Buttons.Item(3).Enabled = False
            Toolbar.Buttons.Item(5).Enabled = False
        End If
        
    Case 7:
        frmSettings.Show vbModal
    
    Case 8:
        frmAbout.Show vbModal
        
    End Select

End Sub

