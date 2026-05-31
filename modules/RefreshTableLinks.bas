Attribute VB_Name = "RefreshTableLinks"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                   RefreshTableLinks                          '
'                                                              '
'      This module contains functions that refresh the         '
'      links to arsh tables if they aren't available.     '
'                                                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit           ' Require variables to be declared before being used.
Option Compare Database   ' Use database order for string comparisons.

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Boolean
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Boolean
Dim kfind As String
Type MSA_OPENFILENAME
    ' Filter string used for the File Open dialog filters.
    ' Use MSA_CreateFilterString() to create this.
    ' Default = All Files, *.*
    strFilter As String
    ' Initial Filter to display.
    ' Default = 1.
    lngFilterIndex As Long
    ' Initial directory for the dialog to open in.
    ' Default = Current working directory.
    strInitialDir As String
    ' Initial file name to populate the dialog with.
    ' Default = "".
    strInitialFile As String
    strDialogTitle As String
    ' Default extension to append to file if user didn't specify one.
    ' Default = System Values (Open File, Save File).
    strDefaultExtension As String
    ' Flags (see constant list) to be used.
    ' Default = no flags.
    lngFlags As Long
    ' Full path of file picked.  On OpenFile, if the user picks a
    ' nonexistent file, only the text in the "File Name" box is returned.
    strFullPathReturned As String
    ' File name of file picked.
    strFileNameReturned As String
    ' Offset in full path (strFullPathReturned) where the file name
    ' (strFileNameReturned) begins.
    intFileOffset As Integer
    ' Offset in full path (strFullPathReturned) where the file extension begins.
    intFileExtension As Integer
End Type

Const ALLFILES = "All Files"

Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As Long
    nMaxCustrFilter As Long
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
    lCustrData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Const OFN_ALLOWMULTISELECT = &H200
Const OFN_CREATEPROMPT = &H2000
Const OFN_EXPLORER = &H80000
Const OFN_FILEMUSTEXIST = &H1000
Const OFN_HIDEREADONLY = &H4
Const OFN_NOCHANGEDIR = &H8
Const OFN_NODEREFERENCELINKS = &H100000
Const OFN_NONETWORKBUTTON = &H20000
Const OFN_NOREADONLYRETURN = &H8000
Const OFN_NOVALIDATE = &H100
Const OFN_OVERWRITEPROMPT = &H2
Const OFN_PATHMUSTEXIST = &H800
Const OFN_READONLY = &H1
Const OFN_SHOWHELP = &H10

Function Findarsh(strSearchPath) As String
' Displays the open file dialog box for the user to locate
' the arsh database. Returns the full path to arsh.

    Dim msaof As MSA_OPENFILENAME

    ' Set options for the dialog box.
    msaof.strDialogTitle = kfind
    msaof.strInitialDir = strSearchPath
    msaof.strFilter = MSA_CreateFilterString("فايل ساعت كارت زني", "*.dbf;*.XLS")

    ' Call the Open File dialog routine.
    MSA_GetOpenFileName msaof

    ' Return the path and file name.
    Findarsh = Trim(msaof.strFullPathReturned)

End Function


Function MSA_CreateFilterString(ParamArray varFilt() As Variant) As String
' Creates a filter string from the passed in arguments.
' Returns "" if no args are passed in.
' Expects an even NUMBER of args (filter name, extension), but
' if an odd NUMBER is passed in, it appends *.*

    Dim strFilter As String
    Dim intRet As Integer
    Dim intNum As Integer

    intNum = UBound(varFilt)
    If (intNum <> -1) Then
        For intRet = 0 To intNum
            strFilter = strFilter & varFilt(intRet) & Chr$(0)
        Next
        If intNum Mod 2 = 0 Then
            strFilter = strFilter & "*.*" & Chr$(0)
        End If

        strFilter = strFilter & Chr$(0)
    Else
        strFilter = ""
    End If

    MSA_CreateFilterString = strFilter
End Function

Function MSA_ConvertFilterString(strFilterIn As String) As String
' Creates a filter string from a bar ("|") separated string.
' The string should pairs of filter|extension strings, i.e. "Access Databases|*.mdb|All Files|*.*"
' If no extensions exists for the last filter pair, *.* is added.
' This code will ignore any empty strings, i.e. "||" pairs.
' Returns "" if the strings passed in is empty.


    Dim strFilter As String
    Dim intNum As Integer, intPos As Integer, intLastPos As Integer

    strFilter = ""
    intNum = 0
    intPos = 1
    intLastPos = 1

    ' Add strings as long as we find bars.
    ' Ignore any empty strings (not allowed).
    Do
        intPos = InStr(intLastPos, strFilterIn, "|")
        If (intPos > intLastPos) Then
            strFilter = strFilter & MID$(strFilterIn, intLastPos, intPos - intLastPos) & Chr$(0)
            intNum = intNum + 1
            intLastPos = intPos + 1
        ElseIf (intPos = intLastPos) Then
            intLastPos = intPos + 1
        End If
    Loop Until (intPos = 0)

    ' Get last string if it exists (assuming strFilterIn was not bar terminated).
    intPos = Len(strFilterIn)
    If (intPos >= intLastPos) Then
        strFilter = strFilter & MID$(strFilterIn, intLastPos, intPos - intLastPos + 1) & Chr$(0)
        intNum = intNum + 1
    End If

    ' Add *.* if there's no extension for the last string.
    If intNum Mod 2 = 1 Then
        strFilter = strFilter & "*.*" & Chr$(0)
    End If

    ' Add terminating NULL if we have any filter.
    If strFilter <> "" Then
        strFilter = strFilter '& Chr$(0)
    End If

    MSA_ConvertFilterString = strFilter
End Function

Private Function MSA_GetSaveFileName(msaof As MSA_OPENFILENAME) As Integer
' Opens the file save dialog.

    Dim of As OPENFILENAME
    Dim intRet As Integer

    MSAOF_to_OF msaof, of
    of.Flags = of.Flags Or OFN_HIDEREADONLY
    intRet = GetSaveFileName(of)
    If intRet Then
        OF_to_MSAOF of, msaof
    End If
    MSA_GetSaveFileName = intRet
End Function

Function MSA_SimpleGetSaveFileName() As String
' Opens the file save dialog with default values.

    Dim msaof As MSA_OPENFILENAME
    Dim intRet As Integer
    Dim strRet As String

    intRet = MSA_GetSaveFileName(msaof)
    If intRet Then
        strRet = msaof.strFullPathReturned
    End If

    MSA_SimpleGetSaveFileName = strRet
End Function

Private Function MSA_GetOpenFileName(msaof As MSA_OPENFILENAME) As Integer
' Opens the file open dialog.

    Dim of As OPENFILENAME
    Dim intRet As Integer

    MSAOF_to_OF msaof, of
    intRet = GetOpenFileName(of)
    If intRet Then
        OF_to_MSAOF of, msaof
    End If
    MSA_GetOpenFileName = intRet
End Function

Function MSA_SimpleGetOpenFileName() As String
' Opens the file open dialog with default values.

    Dim msaof As MSA_OPENFILENAME
    Dim intRet As Integer
    Dim strRet As String

    intRet = MSA_GetOpenFileName(msaof)
    If intRet Then
        strRet = msaof.strFullPathReturned
    End If

    MSA_SimpleGetOpenFileName = strRet
End Function

Public Function CheckLinks() As Boolean
' Check links to the arsh database; returns true if links are OK.
                       On Error Resume Next

    Dim dbs As Database, rst As New ADODB.Recordset

    Set dbs = CurrentDb()

    ' Open linked table to see if connection information is correct.
    Set rst = dbs.OpenRecordset("DEED_HED")

    ' If there's no error, return True.
    If Err = 0 Then
        CheckLinks = True
    Else
        CheckLinks = False
    End If
    kfind = "NEGIN_DB  كجاست؟"
End Function

Public Function CheckLinkpic() As Boolean
' Check links to the arsh database; returns true if links are OK.

    Dim dbs As Database, rst As New ADODB.Recordset

    Set dbs = CurrentDb()

    ' Open linked table to see if connection information is correct.
                       On Error Resume Next
    Set rst = dbs.OpenRecordset("pictur")

    ' If there's no error, return True.
    If Err = 0 Then
        CheckLinkpic = True
    Else
        CheckLinkpic = False
    End If
    kfind = "Pictur  كجاست؟"
End Function

Private Sub OF_to_MSAOF(of As OPENFILENAME, msaof As MSA_OPENFILENAME)
' This sub converts from the win32 structure to the friendly MSAccess structure.

    msaof.strFullPathReturned = Left$(of.lpstrFile, InStr(of.lpstrFile, Chr$(0)))
    msaof.strFileNameReturned = of.lpstrFileTitle
    msaof.intFileOffset = of.nFileOffset
    msaof.intFileExtension = of.nFileExtension
End Sub


Private Sub MSAOF_to_OF(msaof As MSA_OPENFILENAME, of As OPENFILENAME)
' This sub converts from the friendly MSAccess structure to the win32 structure.

    Dim strFile As String * 512

    ' Initialize some parts of the structure.
    of.hwndOwner = Application.hWndAccessApp
    of.hInstance = 0
    of.lpstrCustomFilter = 0
    of.nMaxCustrFilter = 0
    of.lpfnHook = 0
    of.lpTemplateName = 0
    of.lCustrData = 0

    If msaof.strFilter = "" Then
        of.lpstrFilter = MSA_CreateFilterString(ALLFILES)
    Else
        of.lpstrFilter = msaof.strFilter
    End If
    of.nFilterIndex = msaof.lngFilterIndex

    of.lpstrFile = msaof.strInitialFile & String$(512 - Len(msaof.strInitialFile), 0)
    of.nMaxFile = 511

    of.lpstrFileTitle = String$(512, 0)
    of.nMaxFileTitle = 511

    of.lpstrTitle = msaof.strDialogTitle

    of.lpstrInitialDir = msaof.strInitialDir

    of.lpstrDefExt = msaof.strDefaultExtension

    of.Flags = msaof.lngFlags

    of.lStructSize = Len(of)
End Sub

Private Function RefreshLinks(strFileName As String) As Boolean
' Refresh links to the supplied database. Return True if successful.

    Dim dbs As Database
    Dim intCount As Integer
    Dim tdf As TableDef

    ' Loop through all tables in the database.
    Set dbs = CurrentDb()
    For intCount = 0 To dbs.TableDefs.COUNT - 1
       Set tdf = dbs.TableDefs(intCount)
       If tdf.NAME <> "cfg" Then
        ' If the table has a connect string, it's a linked table.
        If Len(tdf.connect) > 0 Then
            tdf.connect = ";DATABASE=" & strFileName
            Err = 0
                               On Error Resume Next
            tdf.RefreshLink         ' Relink the table.
            If Err <> 0 Then
                RefreshLinks = False
                Exit Function
            End If
        End If
      End If
    Next intCount

    RefreshLinks = True        ' Relinking complete.

End Function

Private Function RefreshLinkpic(strFileName As String) As Boolean
' Refresh links to the supplied database. Return True if successful.

    Dim dbs As Database
    Dim intCount As Integer
    Dim tdf As TableDef

    ' Loop through all tables in the database.
    Set dbs = CurrentDb()
       Set tdf = dbs.TableDefs("pictur")
        ' If the table has a connect string, it's a linked table.
        If Len(tdf.connect) > 0 Then
            tdf.connect = ";DATABASE=" & strFileName
            Err = 0
                               On Error Resume Next
            tdf.RefreshLink         ' Relink the table.
            If Err <> 0 Then
                RefreshLinkpic = False
                Exit Function
            End If
        End If

    RefreshLinkpic = True        ' Relinking complete.

End Function

Public Function RelinkTables() As Boolean
' Tries to refresh the links to the arsh database.
' Returns True if successful.

    Const conMaxTables = 8
    Const conNonExistentTable = 3011
    Const conNotarsh = 3078
    Const conNwindNotFound = 3024
    Const conAccessDenied = 3051
    Const conReadOnlyDatabase = 3027
    Const conAppTitle = "نگين  98"

    Dim strAccDir As String
    Dim strSearchPath As String
    Dim strFileName As String
    Dim intError As Integer
    Dim strError As String
    Dim i As Long

    ' Get name of directory where Msaccess.exe is located.
    strAccDir = SysCmd(acSysCmdAccessDir)

    ' Get the default sample database path.
    If Dir(strAccDir & "Samples\.") = "" Then
        strSearchPath = strAccDir
    Else
        strSearchPath = strAccDir & "Samples\"
    End If

    ' Look for the arsh database.
    If (Dir(strSearchPath & "NEGIN_DB.mdb") <> "") Then
        strFileName = strSearchPath & "NEGIN_DB.mdb"
    Else
        ' Can't find arsh, so display the File Open dialog.
      DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "نمي توانم فايل NEGIN_DB را پيدا كنم. شما بايد فايل NEGIN_DB را پيدا كنيد به منظور استفاده" & conAppTitle & "."
        DoCmd.MoveSize 3000, 5000, 7000, 2500
        strFileName = Findarsh(strSearchPath)
        If strFileName = "" Then
            strError = "متاسفم شما بايد فايل NEGIN_DB پيداكنيد براي اجراي" & conAppTitle & "."
            GoTo Exit_Failed
        End If
    End If

    ' Fix the links.
    If RefreshLinks(strFileName) Then   ' It worked!
        RelinkTables = True
        Exit Function
    End If

    ' If it failed, display an error.
    Select Case Err
    Case conNonExistentTable, conNotarsh
        strError = "فايل '" & strFileName & "' داراي جداول نگين  98 نمي باشد."
    Case Err = conNwindNotFound
        strError = "شما نمي توانيد " & conAppTitle & " را اجراء كنيد تا پايگاه نگين  98 را پيدا نكنيد."
    Case Err = conAccessDenied
        strError = "نميتوانم فايل " & strFileName & " را باز كنم زيرا اين فايل فقط خواندني است"
    Case Err = conReadOnlyDatabase
        strError = "نميتوانم جداول  " & conAppTitle & " را مجدداً الصاق كنم زيرا اين فايل فقط خواندني."
    Case Else
        strError = Err.Description
    End Select

Exit_Failed:
   DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, strError
    RelinkTables = False

End Function

Public Function RelinkTablepic() As Boolean
' Tries to refresh the links to the arsh database.
' Returns True if successful.

    Const conMaxTables = 8
    Const conNonExistentTable = 3011
    Const conNotarsh = 3078
    Const conNwindNotFound = 3024
    Const conAccessDenied = 3051
    Const conReadOnlyDatabase = 3027
    Const conAppTitle = "نگين  98"

    Dim strAccDir As String
    Dim strSearchPath As String
    Dim strFileName As String
    Dim intError As Integer
    Dim strError As String
    Dim i As Long

    ' Get name of directory where Msaccess.exe is located.
    strAccDir = SysCmd(acSysCmdAccessDir)

    ' Get the default sample database path.
    If Dir(strAccDir & "Samples\.") = "" Then
        strSearchPath = strAccDir
    Else
        strSearchPath = strAccDir & "Samples\"
    End If

    ' Look for the arsh database.
    If (Dir(strSearchPath & "pictur.mdb") <> "") Then
        strFileName = strSearchPath & "pictur.mdb"
    Else
        ' Can't find arsh, so display the File Open dialog.
         DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, " نمي توانم فايل pictur را پيدا كنم. شما بايد فايل pictur را پيدا كنيد به منظور استفاده" & conAppTitle & "."
         DoCmd.MoveSize 3000, 5000, 7000, 2500
        strFileName = Findarsh(strSearchPath)
        If strFileName = "" Then
            strError = "متاسفم شما بايد فايل pictur پيداكنيد براي اجراي" & conAppTitle & "."
            GoTo Exit_Failed
        End If
    End If

    ' Fix the links.
    If RefreshLinkpic(strFileName) Then   ' It worked!
        RelinkTablepic = True
        Exit Function
    End If

    ' If it failed, display an error.
    Select Case Err
    Case conNonExistentTable, conNotarsh
        strError = "فايل '" & strFileName & "' داراي تصاويرنگين  98 نمي باشد."
    Case Err = conNwindNotFound
        strError = "شما نمي توانيد " & conAppTitle & " را اجراء كنيد تا پايگاه تصوير را پيدا نكنيد."
    Case Err = conAccessDenied
        strError = "نميتوانم فايل " & strFileName & " را باز كنم زيرا اين فايل فقط خواندني است"
    Case Err = conReadOnlyDatabase
        strError = "نميتوانم جداول  " & conAppTitle & " را مجدداً الصاق كنم زيرا اين فايل فقط خواندني."
    Case Else
        strError = Err.Description
    End Select

Exit_Failed:
    DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, strError

    RelinkTablepic = False

End Function

Public Function start()
    If CheckLinks() = False Then
        ' Try to relink the tables; if it fails, shut down.
        If RelinkTables() = False Then
            CloseCurrentDatabase
        End If
    End If
    'If CheckLinkpic() = False Then
      ' Try to relink the tables; if it fails, shut down.
    '    If RelinkTablepic() = False Then
    '        CloseCurrentDatabase
    '    End If
    'End If

End Function
