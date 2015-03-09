VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000                         '  new look commdlg
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
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

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Sub Command1_Click()
    Dim tOPENFILENAME As OPENFILENAME
    Dim lResult As Long
    Dim vFiles As Variant
    Dim lIndex As Long, lStart As Long
    
    With tOPENFILENAME
        .flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES
        .hwndOwner = hWnd
        .nMaxFile = 2048
        .lpstrFilter = "All Files" & Chr(0) & "*.*" & Chr(0) & Chr(0)
        .lpstrFile = Space(.nMaxFile - 1) & Chr(0)
        .lStructSize = Len(tOPENFILENAME)
    End With
    
    lResult = GetOpenFileName(tOPENFILENAME)
    
    If lResult > 0 Then
        With tOPENFILENAME
            vFiles = Split(Left(.lpstrFile, InStr(.lpstrFile, Chr(0) & Chr(0)) - 1), Chr(0))
        End With
        
        If UBound(vFiles) = 0 Then
            List1.AddItem vFiles(0)
        Else
            For lIndex = 1 To UBound(vFiles)
                List1.AddItem AddBS(vFiles(0)) & vFiles(lIndex)
            Next
        End If
    End If
End Sub

Private Function AddBS(ByVal sPath As String) As String
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    AddBS = sPath
End Function
