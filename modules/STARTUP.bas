Attribute VB_Name = "Startup"
Option Compare Database
Option Explicit
Public Enum LocalNwind

    lnNoServer = 0
    lnServerNoDB = 1
    lnIsNorthwind = 2
End Enum
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const NwindAccess = "NorthwindCS"  'the database installed by Access script
Public Const StartupFormName = "Startup"  'this form
Public Const SelectPrompt1 = "Microsoft Access could not find either Microsoft SQL Server 2000 Desktop Engine or Microsoft SQL Server installed on your computer.  Click OK to display the Data Link Properties dialog box and log on to a remote server."
Public Const SelectPrompt2 = "If you want to install SQL Server 2000 Desktop Engine on this computer first, click Cancel, install it from the MSDE2000 folder on the Office XP CD-ROM, and then reopen this Access project."
Public Const SelectTitle = "Select a SQL Server"
Public Const InstallPrompt1 = "Connected to SQL Server on : "
Public Const InstallPrompt2 = "Would you like to install the NorthwindCS sample database?"
Public Const InstallPrompt3 = " This may take a few minutes."
Public Const InstallTitle = "Install Database"
Public Const CreateSuccessful = "Created NorthwindCS database on SQL Server : "
Public Const CreateTitle = "Installation successful"
Public Const CannotFindFile1 = "Cannot find the file :"
Public Const CannotFindFile2 = "Please reinstall the sample project."
Public Const CannotFindTitle = "File not found"
Public Const UnconnectedPrompt = "Set Project to unconnected?"
Public Const fUnconnectedPrompt = False
Public Const OldPhotoDef = "This version of the database has an Employees table with an incorrect data type for the Photo column.  The employee samples that use the Photo column may not work.  To correct this problem, you can either use the table designer to change the data type of the photo column in the Employees table to VARCHAR(255), or reinstall the sample NorthwindCS database."
Dim srvrName As String
Dim msofficeInstance As Boolean

Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNUMBER As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Dim info As OSVERSIONINFO
Dim Ret As Integer        ' OS Information
    Dim ver_major As Integer  ' OS Version
    Dim ver_minor As Integer  ' Minor Os Version
    Dim Build As Long         ' OS Build


Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
                 (LpVersionInformation As OSVERSIONINFO) As Long

Declare Function GetComputerName _
       Lib "kernel32" Alias "GetComputerNameA" ( _
       ByVal lpBuffer As String, nSize As Long) As Long

Private Const MAX_COMPUTERNAME_LENGTH As Long = 15&

' Functions in this module are used in the Startup form.

Function OpenStartup() As Boolean
Dim Msg1 As String
' Displays Startup form
' Used in OnOpen property of Startup form.
 On Error GoTo OpenStartup_Err
    Dim Response, MSG, TITLE, Style
    ' Set the value of HideStartupForm check box using the value of
    ' StartupForm property of database (as set in Display Form box
    ' in Startup dialog box).

    If InStr(1, CurrentProject.Properties("StartupForm"), StartupFormName) > 0 Then
        ' StartupForm property is set to Startup, so clear HideStartupForm
        ' check box.
        Forms(StartupFormName).HideStartupForm = False
    Else
        ' StartupForm property is not set to Startup, so check HideStartupForm
        ' checkbox.
        Forms(StartupFormName).HideStartupForm = True
    End If
    'Get server name
    srvrName = CurrentMachineName
    'srvrName = "(local)"

    'Hide custom menu bar until we're done with the Startup form code, or user sees double
    'DoCmd.ShowToolbar "NorthwindCustomMenuBar", acToolbarNo

    'Check if we're already connected
    msofficeInstance = True
    If CurrentProject.IsConnected Then
        'need code to set project to unconnected.
        Response = vbNo
        MSG = UnconnectedPrompt
        Style = vbYesNo + vbQuestion + vbDefaultButton1
        TITLE = ""
        If (fUnconnectedPrompt) Then
            Response = MsgBox(MSG, Style, TITLE)
        End If
        If Response = vbYes Then
            'write blank string for testing
            CurrentProject.OpenConnection "Provider="
        Else
            'check if we are connected to the expected db
            CheckConnectedServer
        End If

    Else
       'attempt to autostart server
        If (StartSQLServer) Then
            CheckConnectedServer
        Else
            MSG = SelectPrompt1 + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13) + SelectPrompt2
            Style = vbOKCancel + vbInformation
            TITLE = SelectTitle
            Response = MsgBox(MSG, Style, TITLE)
            If Response <> vbCancel Then
                DoCmd.RunCommand acCmdConnection
                If CurrentProject.IsConnected Then CheckConnectedServer
            End If
         End If
    End If


OpenStartup_Exit:
    'If we're not connected, hide the HideStartupForm checkbox and label so that the user
    'won't disable the startup code from running next time
    If CurrentProject.BaseConnectionString = "" Then
        Forms(StartupFormName).HideStartupForm.Visible = False
        Forms(StartupFormName).HideStartupFormLabel.Visible = False
    End If

    Exit Function


OpenStartup_Err:
    Const conPropertyNotFound = 3270
    If Err = conPropertyNotFound Then
        Forms(StartupFormName).HideStartupForm = True
        Resume OpenStartup_Exit
    End If
    MsgBox Err.Description
    GoTo OpenStartup_Exit

End Function

Function CreateDB() As Boolean
 On Error GoTo CreateDB_Err
    Dim cn As New ADODB.Connection, cm As New ADODB.Command
    Dim devno As Integer, devpath As String
    Dim ver As Integer
    'Open up a new connection direct to SQL Server (not through MSDataShape)
    cn.connectionString = CurrentProject.BaseConnectionString
    cn.Open

    cm.ActiveConnection = cn
    cm.CommandType = adCmdText

    cm.CommandText = "use master"
    cm.Execute

    ver = CInt(MID(cn.Properties("DBMS Version"), 2, 1))
    If ver >= 7 Then
        cm.CommandText = "create database " + NwindAccess
        cm.Execute
    Else
        If GetDevInfo(cm, devno, devpath) Then
            cm.CommandText = "DISK INIT NAME = '" + NwindAccess + _
                "',PHYSNAME = '" + devpath + NwindAccess + _
                ".dat', VDEVNO = " + CStr(devno) + ", Size = 2560"
            cm.Execute
            cm.CommandText = "CREATE DATABASE " + NwindAccess + " ON " + NwindAccess + " = 5"
            cm.Execute
         End If
    End If
    cm.CommandText = "exec sp_dboption '" + NwindAccess + "','trunc. log on chkpt.','true'"
    cm.Execute
    cm.CommandText = "exec sp_dboption '" + NwindAccess + "','trunc. log on chkpt.','true'"
    cm.Execute

    CreateDB = True
    cn.Close
    Exit Function

CreateDB_Err:
    MsgBox Err.Description
    CreateDB = False
    cn.Close
End Function
Function RunScript(dbname As String, InFile As String, VERN As String, CHK As Boolean) As Boolean
                On Error Resume Next
    Dim cn As New ADODB.Connection, cm As New ADODB.Command, dtt1, dtt2 As Long
    Dim infileopen As Boolean
    Dim incmd$, batch$, Q$
    Q$ = """"  'quoted id char
    'Open up a new connection direct to SQL Server
    cn.connectionString = CurrentProject.BaseConnectionString
    cn.Open
    cm.ActiveConnection = cn
    cm.CommandType = adCmdText
    cm.CommandText = "use " + Q$ + dbname + Q$
    cm.Execute
    Err.Clear
    Dim ab As Boolean
    Open InFile For Input As #1
    If Err = 0 Then
        If IsLoaded("About") Then
            ab = True
        Else
            ab = False
        End If
        infileopen = True
        If CHK Then
            Line Input #1, incmd$
            dtt1 = MID(VERN, 24, 2) * 10000 + MID(VERN, 27, 2) * 100 + MID(VERN, 30, 2)
            dtt2 = MID(incmd$, 24, 2) * 10000 + MID(incmd$, 27, 2) * 100 + MID(incmd$, 30, 2)
        End If
        If dtt2 >= dtt1 Or Not CHK Then
            Do While Not EOF(1)
                Line Input #1, incmd$
                If (Left(Trim(incmd$), 2) <> "go" Or Left(Trim(incmd$), 2) <> "GO") Then
                    batch$ = batch$ + " " + incmd$ + Chr(10) + Chr(13)
                Else
                    If batch$ <> "" Then
                        cm.CommandText = batch$
                       ' ' DEBUG.PRINT cm.CommandText
                        cm.Execute
                        batch$ = ""
                    End If
                End If
                If ab Then
                    Forms![About]![IDD] = Forms![About]![IDD] + 1
                    Forms![About].Form.Repaint
                Else
                    Forms![Startup]![IDD] = Forms![Startup]![IDD] + 1
                    Forms![Startup].Form.Repaint
                End If
'                Forms![bun].Form.Refresh
            Loop
        Else
             DoCmd.OpenForm "MESAG", , , , , acDialog, "ورژن اسكريبت**NEGIN.SCR**: " & incmd$ & "*** با ورژن نرم افزار: " & VERN & "**** تطبيق ندارد و قابل اجرا نيست بايد از سايت دانلود نماييد"
        End If
        Close #1
        RunScript = True
        cn.Close
        Call SCRIPTIN
    Else
        Call SCRIPTIN
    End If
End Function
Function HideStartupForm()

'End Function
 On Error GoTo HideStartupForm_Err
' Uses the value of HideStartupForm check box to determine the setting for
' StartupForm property of database. (The setting is displayed in Display Form
' box in Startup dialog box).
' Used in OnClick event of HideStartupForm checkbox of Startup form.
        If Forms(StartupFormName).HideStartupForm Then
                CurrentProject.Properties("StartupForm") = "Main Switchboard"
        Else
            ' HideStartupForm check box is cleared, so set StartupForm property to Startup.
            CurrentProject.Properties("StartupForm") = StartupFormName
        End If

        Exit Function

HideStartupForm_Err:
    Const conPropertyNotFound = 2473
    If Err = conPropertyNotFound Then
        CurrentProject.Properties.ADD "StartupForm", StartupFormName
        Resume Next
    End If
End Function

Function CloseForm()
' Closes Startup form.
' Used in OnClick property of OK command button on Startup form.
    DoCmd.Close
    If (CurrentProject.IsConnected) Then
        DoCmd.OpenForm ("Main Switchboard")
    End If
End Function

Function CheckConnectedServer()
'looks for Northwind or NorthwindCS on selected server
 On Error GoTo CheckConnectedServer_Err
    Dim cn As ADODB.Connection
    Dim ServerStr As String, DBStr As String, ConnectStr As String
    Dim Response, MSG, TITLE, Style

    Set cn = CurrentProject.Connection
    DBStr = cn.Properties("Initial Catalog")

    If (DBStr = NwindAccess) Then
        'we're done
    ElseIf CheckForNorthwind(cn) Then
        'we're done
    Else
        'No northwind database on this server, ask if they want to install one
        ServerStr = cn.Properties("Data Source")
        If (DBInstallPrompt(ServerStr)) Then
            MSG = CreateSuccessful + CurrentProject.Connection.Properties("Data Source")
            Style = vbOKOnly + vbInformation
            TITLE = CreateTitle
            MsgBox MSG, Style, TITLE
            ChangeDB (NwindAccess)
         Else
            'set project back to unconnected
            CurrentProject.OpenConnection "Provider="
         End If
    End If

    Exit Function

CheckConnectedServer_Err:
    MsgBox Err.Description

End Function


Function DBInstallPrompt(server As String) As Boolean
 On Error GoTo DBInstallPrompt_Err
    Dim Response, MSG, TITLE, Style
    Dim serverName As String, Script As String
    DBInstallPrompt = False

    If (server = "(local)") Then serverName = "this machine" Else serverName = server
    'Check for script file before offering to install
    Script = CurrentProject.PATH + "\" + NwindAccess + ".sql"
    If Dir(Script) = "" Then
        MSG = CannotFindFile1 + Script + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13) + CannotFindFile2
        Style = vbOKOnly + vbExclamation
        TITLE = CannotFindTitle
        Response = MsgBox(MSG, Style, TITLE)
        Exit Function
    Else
        MSG = InstallPrompt1 + CurrentProject.Connection.Properties("Data Source") + _
                Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13) + InstallPrompt2 + _
                Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13) + InstallPrompt3
        Style = vbYesNo + vbQuestion + vbDefaultButton1
        TITLE = InstallTitle
        Response = MsgBox(MSG, Style, TITLE)
        If Response = vbNo Then
            DBInstallPrompt = False
            Exit Function
        End If
    End If

    Screen.MousePointer = 11
    If CreateDB() Then
        If RunScript(NwindAccess, Script, "1", False) Then
           DBInstallPrompt = True
        End If
    End If
    Screen.MousePointer = 0

    Exit Function

DBInstallPrompt_Err:
    Screen.MousePointer = 0
    MsgBox Err.Description
    DBInstallPrompt = False
    Exit Function

End Function



Public Function GetDevInfo(cm As ADODB.Command, FreeDevNo As Integer, MasterPath As String) As Boolean
 On Error GoTo GetDevInfo_Err
    Dim rs As New ADODB.Recordset
    Const startdev As Integer = 10
    Dim i As Integer
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    ' need to set ansi_defaults off for next stmt, then reset for normal OLE DB operation
    cm.CommandText = "SET ANSI_DEFAULTS OFF exec sp_helpdevice SET ANSI_DEFAULTS ON SET CURSOR_CLOSE_ON_COMMIT OFF SET IMPLICIT_TRANSACTIONS OFF"
    rs.Open cm
    rs.Sort = "device_NUMBER"
    'look for lowest devno above start point that is not in use
    i = startdev
    Do Until rs.EOF
        rs.MoveNext
        If rs.Fields("device_NUMBER") = i Then i = i + 1
        If rs.Fields("device_NUMBER") > i Then Exit Do
    Loop
    If rs.EOF Then
        MsgBox "Couldn't get free device NUMBER"
        GetDevInfo = False
        Exit Function
    End If
    FreeDevNo = i

    rs.Filter = "device_name = 'master'"
    i = InStr(1, rs.Fields("physical_name"), "master.dat")
    If (i = 0) Then
        MsgBox "Couldn't find master device path."
        GetDevInfo = False
        Exit Function
    End If

    MasterPath = Left(rs.Fields("physical_name"), i - 1)

    GetDevInfo = True
    rs.Close
    Exit Function

GetDevInfo_Err:
    MsgBox Err.Description
    GetDevInfo = False
    Exit Function


End Function

Public Function CheckForNorthwind(cn As ADODB.Connection) As Boolean
 On Error GoTo CheckForNorthwind_Err

    Dim rs As New ADODB.Recordset
    Set rs = cn.Execute("select 1 from master..sysdatabases where name='" + NwindAccess + "'")
    If rs.EOF Then
        rs.Close
        CheckForNorthwind = False
    Else
        'found the Access version of the database
        rs.Close
        ChangeDB (NwindAccess)
        CheckForNorthwind = True
    End If

Exit Function

CheckForNorthwind_Err:
    MsgBox Err.Description
    CheckForNorthwind = False
End Function
Public Function StartSQLServer() As Boolean
 On Error GoTo StartSQLServer_Err
    StartSQLServer = False
    Dim MSG As String
    Dim srv As Object
    Dim t As Integer, ET As Integer
    Dim fNewADP As Boolean, DataSource As String, ConnectStr As String, srvstr As String

    Dim WaitForStarting As Integer, WaitForRunning As Integer, WaitAfterRunning As Integer
    WaitForStarting = 10 'wait 10 seconds after requesting server start to get status=starting
    WaitForRunning = 20 'wait 20 seconds after getting status=starting to get status=running
    WaitAfterRunning = 5  'wait 5 seconds after status=running before attempting to connect
    Dim StatusMsg As String
    StatusMsg = "Attempting to start server "

connect:
    ConnectStr = CurrentProject.BaseConnectionString
    If ConnectStr = "" Then
        fNewADP = True
        If (msofficeInstance = True) Then
            DataSource = srvrName + "\MSOFFICE"
        Else
            DataSource = srvrName
        End If
    Else
        fNewADP = False
        Dim cn As New ADODB.Connection
        cn.connectionString = ConnectStr
        DataSource = cn.Properties("Data Source")

    End If

    Set srv = CreateObject("SQLDMO.SQLServer")
    srv.NAME = DataSource

    If srv.STATUS = 1 Then 'SQLDMOSvc_Running
        'if server already running don't try to start
        GoTo ServerRunning
    ElseIf srv.STATUS <> 3 Then 'SQLDMOSvc_Stopped
        'if server isn't in a stopped state don't try to start
        ' DEBUG.PRINT "SQL Server in unexpected state :  " + CStr(srv.STATUS)
        GoTo StartSQLServer_Exit
    End If

    ET = WaitForStarting + WaitForRunning + WaitAfterRunning
    SysCmd acSysCmdInitMeter, StatusMsg, ET
    ET = 0

    srv.start False
    'first loop until server is starting
    t = 0
    Do Until srv.STATUS = 4  'SQLDMOSvc_Starting
        If (t = WaitForStarting) Then
            ' DEBUG.PRINT "timed out waiting for status=starting"
            GoTo StartSQLServer_Exit
        End If
        Sleep (1000)
        t = t + 1
        ET = ET + 1
        SysCmd acSysCmdUpdateMeter, ET
    Loop
    'now loop until running or stopped
    t = 0
    Do Until srv.STATUS = 1  'SQLDMOSvc_Running
        If (t = WaitForRunning Or srv.STATUS = 3) Then
            ' DEBUG.PRINT "timed out waiting for status=running"
            GoTo StartSQLServer_Exit
        End If
        Sleep (1000)
        t = t + 1
        ET = ET + 1
        SysCmd acSysCmdUpdateMeter, ET
    Loop

    'Now wait a few seconds before attempting to connect
    t = 0
    Do Until t = WaitAfterRunning
        Sleep (1000)
        t = t + 1
        ET = ET + 1
        SysCmd acSysCmdUpdateMeter, ET
    Loop

ServerRunning:
    If fNewADP Then
        'get the local server's network name
        If (msofficeInstance = True) Then
            srvstr = srvrName + "\MSOFFICE"
        Else
            srvstr = "(local)"
        End If
            info = GetOSVersion
            If (info.dwPlatformId >= 2) Then
                srv.LoginSecure = True
                srv.connect srvstr, "sa", ""
                ConnectStr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=master;" + _
                            "Data Source=" + srvstr
            Else
                srv.connect srvstr, "sa", ""
                ConnectStr = "Provider=SQLOLEDB;Data Source=" + srvstr + _
                        ";User ID=sa;Password=;Persist Security Info=TRUE;Initial Catalog=master"
            End If
    End If

    CurrentProject.OpenConnection ConnectStr

    SysCmd acSysCmdClearStatus
    StartSQLServer = True
    Exit Function

StartSQLServer_Err:
    'Don't show the user errors in attempts to start server
    ' DEBUG.PRINT "Error " + CStr(Err.NUMBER) + " :  " + Err.Description
    If (Err.NUMBER = 1060 Or Err.NUMBER = 1707 Or Err.NUMBER = 123 Or Err.NUMBER = 1722) Then
        If msofficeInstance = False Then
            Resume StartSQLServer_Exit
        Else
            msofficeInstance = False
            Resume connect
        End If
    End If

StartSQLServer_Exit:
    SysCmd acSysCmdClearStatus

End Function


Public Function CurrentMachineName() As String
       Dim lSize As Long
       Dim sBuffer As String
       sBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)
       lSize = Len(sBuffer)

       If GetComputerName(sBuffer, lSize) Then
              CurrentMachineName = Left$(sBuffer, lSize)
       End If
End Function

Public Function GetOSVersion() As OSVERSIONINFO
    Dim verinfo As OSVERSIONINFO
    Dim vers As Integer
          verinfo.dwOSVersionInfoSize = Len(verinfo)
          Ret = GetVersionEx(verinfo)
          GetOSVersion = verinfo
End Function

Public Function IsNorthwindNew()
' This function is used to determine if the NorthwindCS database that the project is
' connected to is using a Varchar data type for the Employees table Photo field.  The
' older version of NorthwindCS used an Image data type and that will no longer work
' with the new Employees form and data page.
 On Error GoTo IsNorthwindNew_Err

    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim fieldName As ADODB.Field
    Dim fieldType As Integer

    Set cn = CurrentProject.Connection
    Set rs = New ADODB.Recordset

    rs.ActiveConnection = cn
    rs.Open "Employees", cn, , , adCmdTable
    fieldType = CInt(rs.Fields.item("Photo").Type)
    If (fieldType <> 200) Then
        MsgBox OldPhotoDef
    End If

IsNorthwindNew_Err:
' Silently ignore any errors that could occur checking the Employees table
    Exit Function
End Function
