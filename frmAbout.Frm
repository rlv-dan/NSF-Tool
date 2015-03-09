VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About NSF Tool"
   ClientHeight    =   2430
   ClientLeft      =   4620
   ClientTop       =   4440
   ClientWidth     =   3255
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2430
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3015
      Begin VB.Image Image1 
         Height          =   525
         Left            =   120
         Picture         =   "frmAbout.frx":000C
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "v2.0"
         Height          =   195
         Left            =   1860
         TabIndex        =   3
         Top             =   560
         Width           =   315
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "http://www.rlvision.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         MouseIcon       =   "frmAbout.frx":0938
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label15 
         Caption         =   "Freeware by RL Vision © 2000-2010"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cool!"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'XP Themes
    If IsThemed() Then
        FixThemeSupport Controls
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call UnloadXpApp

End Sub

Private Sub Label9_Click()

    Temp = ShellExecute(Me.hWnd, vbNullString, "http://www.rlvision.com", vbNullString, "c:\", SW_SHOWNORMAL)

End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    
    SetCursor LoadCursor(0, IDC_HAND)   'set hand cursor

End Sub
