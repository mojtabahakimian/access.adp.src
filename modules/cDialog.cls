VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'
'  Copyright (C)1998-99 Dev Ashish and Terry Kreft, All Rights Reserved
'  The Access Web (http://home.att.net/~dashish)
'  Comments and bug reports can be emailed to us
'  Dev Ashish (dash10@hotmail.com) ; Terry Kreft (terry.kreft@mps.co.uk)
'
'

Private Type OPENFILENAME
  lStructSize As Long
  hWnd As Long
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
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib _
  "comdlg32.dll" Alias "GetOpenFileNameA" _
  (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib _
  "comdlg32.dll" Alias "GetSaveFileNameA" _
  (pOpenfilename As OPENFILENAME) As Long

Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SAVE = 0
Private Const OFN_OPEN = 1

Private lnghWnd As Long
Private wMode As Integer
Private szDialogTitle As String
Private szFileName As String
Private szFilter As String
Private szDefDir As String
Private szDefExt As String
Private szFileTitle As String
Private szFileDir As String
Private intFilterIndex As Integer

'BrowseFolder Declarations
Private Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias _
            "SHGetPathFromIDListA" (ByVal pidl As Long, _
            ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias _
            "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) _
            As Long

Private Const BIF_RETURNONLYFSDIRS = &H1

Public Function BrowseFolder(szDialogTitle As String) As String
  Dim x As Long, BI As BROWSEINFO, dwIList As Long
  Dim szPath As String, wPos As Integer

    With BI
        .hOwner = lnghWnd
        .lpszTitle = szDialogTitle
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    dwIList = SHBrowseForFolder(BI)
    szPath = Space$(512)
    x = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)

    If x Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Left$(szPath, wPos - 1)
    Else
        BrowseFolder = ""
    End If
End Function

Public Function Action() As String
  Dim x As Long, OFN As OPENFILENAME
  Call SetDefs
  With OFN
    .lStructSize = Len(OFN)
    .hWnd = lnghWnd
    .lpstrTitle = szDialogTitle
    .lpstrFile = szFileName & String$(250 - Len(szFileName), 0)
    .nMaxFile = 255
    .lpstrFileTitle = String$(255, 0)
    .nMaxFileTitle = 255
    .lpstrFilter = szFilter
    .nFilterIndex = intFilterIndex
    .lpstrInitialDir = szDefDir
    .lpstrDefExt = szDefExt
    If wMode = 1 Then
      OFN.Flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST _
                  Or OFN_FILEMUSTEXIST
      x = GetOpenFileName(OFN)
    Else
      OFN.Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT _
                  Or OFN_PATHMUSTEXIST Or OFN_CREATEPROMPT
      x = GetSaveFileName(OFN)
    End If
    If x <> 0 Then
      If InStr(.lpstrFile, Chr$(0)) > 0 Then
        szFileName = Left$(.lpstrFile, InStr(.lpstrFile, Chr$(0)) - 1)
        szFileTitle = Left$(.lpstrFileTitle, InStr(.lpstrFileTitle, Chr$(0)) - 1)
        Call getFile_Dir
      End If
    Else
      szFileName = ""
    End If
  End With
  Action = szFileName
End Function
'Pass a bar separated string and returns a Null separated string
Private Function NullSepString(ByVal BarString As String) As String
  Dim intInstr As Integer
  Const vbBar = "|"
  Do
    intInstr = InStr(BarString, vbBar)
    If intInstr > 0 Then Mid(BarString, intInstr, 1) = vbNullChar
  Loop While intInstr > 0
  NullSepString = BarString
End Function
Private Sub getFile_Dir()
  Dim intInstr As Integer
  intInstr = InStr(szFileName, szFileTitle) - 1
  szFileDir = Left(szFileName, intInstr)
End Sub

Property Let hWnd(SourcehWnd As Long)
  lnghWnd = SourcehWnd
End Property

Property Let ModeOpen(DialogMode As Boolean)
  wMode = DialogMode * -1
End Property
Property Let TITLE(DialogTitle As String)
  szDialogTitle = DialogTitle
End Property
Property Let fileName(DefaultFile As String)
  szFileName = DefaultFile
End Property
Property Get fileName() As String
  fileName = szFileName
End Property
Property Let Filter(FilterList As String)
  szFilter = NullSepString(FilterList)
End Property
Property Let StartDir(InitialDir As String)
  szDefDir = InitialDir
End Property
Property Let DefaultExtension(DefExt As String)
  szDefExt = DefExt
End Property
Property Get FileTitle()
  FileTitle = szFileTitle
End Property
Property Get FileDir() As String
  FileDir = szFileDir
End Property
Private Sub SetDefs()
  If lnghWnd = 0 Then lnghWnd = 0
  If szFilter = "" Then szFilter = NullSepString("All Files|*.*")
  If szDefDir = "" Then szDefDir = "C:\"
  If intFilterIndex = 0 Then intFilterIndex = 1
End Sub
