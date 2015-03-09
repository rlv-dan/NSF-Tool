VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File information"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRipper 
      Height          =   285
      Left            =   1320
      MaxLength       =   31
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtCopyright 
      Height          =   285
      Left            =   1320
      MaxLength       =   31
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txtArtist 
      Height          =   285
      Left            =   1320
      MaxLength       =   31
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1320
      MaxLength       =   31
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblFiletype 
      Caption         =   "Label8"
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "Type:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblRipper 
      Caption         =   "Ripper:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2550
      Width           =   1455
   End
   Begin VB.Label lblSpecials 
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label lblSongs 
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label lblFilename 
      Caption         =   "lblFilename"
      Height          =   615
      Left            =   1320
      TabIndex        =   12
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Sound Chip:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   4440
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000011&
      BorderWidth     =   2
      X1              =   120
      X2              =   4440
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   4440
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label5 
      Caption         =   "Artist:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1830
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Copyright:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2190
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Songs:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Title:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1470
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   4440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000011&
      BorderWidth     =   2
      X1              =   120
      X2              =   4440
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      BorderWidth     =   2
      X1              =   120
      X2              =   4440
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myNSF As NSF

Private Sub cmdSave_Click()

    num = frmMain.ListView.SelectedItem.Index

    myNSF.Title = Trim(txtTitle)
    myNSF.Artist = Trim(txtArtist)
    myNSF.Copyright = Trim(txtCopyright)
    If myNSF.NSFE = True Then myNSF.Ripper = Trim(txtRipper)

    If myNSF.NSFE = False Then
        If txtTitle = "?" Or Trim(txtTitle) = "" Then txtTitle = "<?>"
        If txtArtist = "?" Or Trim(txtArtist) = "" Then txtArtist = "<?>"
        If txtCopyright = "?" Or Trim(txtCopyright) = "" Then txtCopyright = "<?>"
    End If

    Dim ret
    If myNSF.NSFE = True Then
        ret = SaveNsfExtended(myNSF)
    Else
        ret = SaveNsfInfo(myNSF)
    End If
    If ret = False Then
        frmMain.ListView.ListItems.Remove (num)
        frmMain.ListView1.ListItems.Remove (num)
        frmMain.Enabled = True
        Unload Me
        Exit Sub
    End If
    
    ret = LoadNSF(myNSF)
    If ret = True Then
        Set itmX = frmMain.ListView.SelectedItem
        itmX.Text = RemovePath(myNSF.Path)
        itmX.SubItems(1) = FormatName(myNSF)
    End If
    
    Unload Me

End Sub

Private Sub cmdOk_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'XP Themes
    If IsThemed() Then
        FixThemeSupport Controls
    End If

    num = frmMain.ListView.SelectedItem.Index
    myNSF.Path = frmMain.ListView1.ListItems.Item(num)
    ret = LoadNSF(myNSF)
    If ret = False Then
        ret = MsgBox("Unable to show info...", vbCritical, "Error")
        frmMain.Enabled = True
        Unload Me
    End If

    lblFilename = RemovePath(myNSF.Path)
        
    txtTitle = myNSF.Title
    txtArtist = myNSF.Artist
    txtCopyright = myNSF.Copyright
    
    If myNSF.NSFE = True Then
        lblFiletype = "Extended NSF"
        txtRipper = myNSF.Ripper
        txtRipper.Visible = True
        lblRipper.Visible = True
    Else
        lblFiletype = "NSF"
        txtRipper.Visible = False
        lblRipper.Visible = False
    End If

    If (myNSF.System And 1) = 0 Then lblFiletype = lblFiletype & " (NTSC)"
    If (myNSF.System And 1) = 1 Then lblFiletype = lblFiletype & " (PAL)"
    If (myNSF.System And 2) = 1 Then lblFiletype = lblFiletype & " (Dual NTSC/PAL)"

    lblSongs = myNSF.Songs
    
    If (myNSF.Specials And 1) = 1 Then lblSpecials = lblSpecials + "Konami VRC1"
    If (myNSF.Specials And 2) = 2 Then lblSpecials = lblSpecials + "Konami VRC2"
    If (myNSF.Specials And 4) = 4 Then lblSpecials = lblSpecials + "FDS Sound"
    If (myNSF.Specials And 8) = 8 Then lblSpecials = lblSpecials + "MMC5 Audio"
    If (myNSF.Specials And 16) = 16 Then lblSpecials = lblSpecials + "Namco 106"
    If (myNSF.Specials And 32) = 32 Then lblSpecials = lblSpecials + "Sunsoft FME-07"
    If myNSF.Specials = 0 Then lblSpecials = "Standard"

    cmdSave.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call UnloadXpApp

End Sub

Private Sub txtArtist_Change()

    cmdSave.Enabled = True
    cmdSave.Default = True

End Sub

Private Sub txtCopyright_Change()

    cmdSave.Enabled = True
    cmdSave.Default = True

End Sub

Private Sub txtTitle_Change()

    cmdSave.Enabled = True
    cmdSave.Default = True

End Sub
