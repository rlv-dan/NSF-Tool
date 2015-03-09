VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReplace 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Text            =   "Text5"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.ComboBox cmbReplace 
      Height          =   315
      ItemData        =   "frmSettings.frx":000C
      Left            =   360
      List            =   "frmSettings.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   4695
      Begin VB.CheckBox chkConfirm 
         Caption         =   "Manually confirm each renaming"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rename Format"
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      Begin VB.TextBox txtFormat 
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label lblInfo 
         Caption         =   "lblInfo"
         Height          =   1455
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame Frame6 
      Caption         =   "Replacements when Renaming"
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   4695
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private replace_string(12) As String

Private Sub Form_Load()

    'XP Themes
    If IsThemed() Then
        FixThemeSupport Controls
    End If
    
    lblInfo = "%T = Title" & Chr(10) & "%A = Artist" & Chr(10) & "%C = Copyright Holder" & Chr(10) & "%S = Number Of Songs" & Chr(10) & "%F = Extra Sound Chip Features" & Chr(10) & "%R = Ripper (NSFE Only)" & Chr(10) & "%V = TV System (NTSC/PAL)"
    '& Chr(10) & "%S = Special Sound Chips" & Chr(10) & "%P = PAL/NTSC/DUAL"
    
    If ConfirmRename = True Then chkConfirm = 1 Else chkConfirm = 0
    
    txtFormat = nsfFormat
    
    cmbReplace.Clear
    cmbReplace.AddItem "Konami VRCVI"
    cmbReplace.AddItem "Konami VRCVII"
    cmbReplace.AddItem "FDS Sound"
    cmbReplace.AddItem "MMC5 Audio"
    cmbReplace.AddItem "Namco 106"
    cmbReplace.AddItem "Sunsoft FME - 7"
    cmbReplace.AddItem "Unknown Title"
    cmbReplace.AddItem "Unknown Artist"
    cmbReplace.AddItem "Unknown Copyright"
    cmbReplace.AddItem "Unknown Ripper"
    cmbReplace.AddItem "NTSC"
    cmbReplace.AddItem "PAL"
    cmbReplace.AddItem "Dual NTSC/PAL"
    
    replace_string(0) = special_bit0
    replace_string(1) = special_bit1
    replace_string(2) = special_bit2
    replace_string(3) = special_bit3
    replace_string(4) = special_bit4
    replace_string(5) = special_bit5
    replace_string(6) = unknown_title
    replace_string(7) = unknown_artist
    replace_string(8) = unknown_copyright
    replace_string(9) = unknown_ripper
    replace_string(10) = ntsc_string
    replace_string(11) = pal_string
    replace_string(12) = ntscpal_string
    cmbReplace.ListIndex = 0


End Sub

Private Sub cmbReplace_Click()
    
        txtReplace.Text = replace_string(cmbReplace.ListIndex)

End Sub

Private Sub cmdOk_Click()

    'save settings
    
    If txtFormat = "" Then
        ret = MsgBox("Not a valid format...", vbCritical, "Error")
        Exit Sub
    End If
    
    test = False    'make sure at least one % is included
    For a = 1 To Len(txtFormat)
        If Mid(txtFormat, a, 1) = "%" Then test = True
    Next
    If test = False Then
        ret = MsgBox("Not a valid format...", vbCritical, "Error")
        Exit Sub
    End If
    
    Temp = txtFormat    'remove invalid chars
    txtFormat = ""
    varning = False
    For num = 1 To Len(Temp)
        ch = Mid(Temp, num, 1)
        If ch <> "\" And ch <> "/" And ch <> ":" And ch <> "*" And ch <> "?" And ch <> "<" And ch <> ">" And ch <> "|" And ch <> Chr(34) Then
            txtFormat = txtFormat + ch
        Else
            varning = True
        End If
    Next
    If varning = True Then
        ret = MsgBox("One or more characters in the replacement strings are invalid in filenames and have been removed...", vbCritical, "Warning")
    End If
    
    'Update replacements
    special_bit0 = replace_string(0)
    special_bit1 = replace_string(1)
    special_bit2 = replace_string(2)
    special_bit3 = replace_string(3)
    special_bit4 = replace_string(4)
    special_bit5 = replace_string(5)
    unknown_title = replace_string(6)
    unknown_artist = replace_string(7)
    unknown_copyright = replace_string(8)
    unknown_ripper = replace_string(9)
    ntsc_string = replace_string(10)
    pal_string = replace_string(11)
    ntscpal_string = replace_string(12)
    ''''''''''''''''''''
    
    nsfFormat = txtFormat
    UpdateListview
    
    If chkConfirm = 1 Then ConfirmRename = True Else ConfirmRename = False
    
    Unload Me

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call UnloadXpApp

End Sub

Private Sub txtReplace_Change()

    replace_string(cmbReplace.ListIndex) = txtReplace.Text

End Sub
