VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "http"
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

Private Type URL_COMPONENTS
  dwStructSize As Long
  lpszScheme As String
  dwSchemeLength As Long
  nScheme As Long
  lpszHostName As String
  dwHostNameLength As Long
  nPort As Long
  lpszUserName As String
  dwUserNameLength As Long
  lpszPassword As String
  dwPasswordLength As Long
  lpszUrlPath As String
  dwUrlPathLength As Long
  lpszExtraInfo As String
  dwExtraInfoLength As Long
End Type

Private mstrURL As String
Private mstrDestination As String
Private mblnConnectState As Boolean
Private mblnOverWrite As Boolean
Private hFile As Long
Private hSession As Long
Private lnghWnd As Long
Private hInet As Long
Private hFTP As Long
Private hURL As Long
Private mlngSize As Long
Private mblnPromptForFile As Boolean
Private mtURLInfo As URL_COMPONENTS
Private mcolRemoteDir As Collection
Private mstrSrcFile As String
Private mblnUpload As Boolean
Private mblnUseProxy As Boolean
Private mblnCreateRemoteDir As Boolean
Private mintErrorTrap As Integer

Private Const mconERR_BAD_URL As Long = vbObjectError + 1000
Private Const mconERR_REQ_FAILED As Long = vbObjectError + 2000
Private Const mconERR_UNKNOWN = vbObjectError + 3000
Private Const mconERR_CONNECTION_FAIL = vbObjectError + 4001
Private Const mconERR_CANNOT_START_TRANSFER = vbObjectError + 5000
Private Const mconERR_WRONG_OPERATION = vbObjectError + 5000

Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H2
Private Const FTP_TRANSFER_TYPE_MASK = FTP_TRANSFER_TYPE_ASCII _
                                    Or FTP_TRANSFER_TYPE_BINARY

Private Const ICU_NO_ENCODE = &H20000000   '// Don't convert unsafe characters to escape sequence
Private Const ICU_DECODE = &H10000000      '//Convert %XX escape sequences to characters
Private Const ICU_NO_META = &H8000000      '//Don't convert .. etc. meta path sequences
Private Const ICU_ENCODE_SPACES_ONLY = &H4000000  '// Encode spaces only
Private Const ICU_BROWSER_MODE = &H2000000
Private Const ICU_ESCAPE = &H80000000      '// (un)escape URL characters

Private Const MAX_CHUNK = 4096
Private Const MAX_PATH = 260
Private Const MAX_BUFFER = 1024

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const CREATE_ALWAYS = 2
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000

Private Const INTERNET_CONNECTION_MODEM = 1
Private Const INTERNET_CONNECTION_LAN = 2
Private Const INTERNET_CONNECTION_PROXY = 4
Private Const INTERNET_CONNECTION_MODEM_BUSY = 8

Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2

Private Const INTERNET_FLAG_ASYNC = &H10000000
Private Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Private Const INTERNET_FLAG_RAW_DATA = &H40000000
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_DONT_CACHE = &H4000000
Private Const INTERNET_FLAG_MAKE_PERSISTENT = &H2000000
Private Const INTERNET_FLAG_PASSIVE = &H8000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const INTERNET_FLAG_TRANSFER_BINARY = &H2

Private Const INTERNET_SCHEME_UNKNOWN = -1

Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0&
Private Const INTERNET_OPEN_TYPE_DIRECT = 1&
Private Const INTERNET_OPEN_TYPE_PROXY = 3&
Private Const INTERNET_DEFAULT_FTP_PORT = 21
Private Const INTERNET_INVALID_PORT_NUMBER = 0
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_SERVICE_GOPHER = 2
Private Const INTERNET_SERVICE_HTTP = 3

Private Const INTERNET_ERROR_BASE = 12000

Private Const ERROR_BAD_PATHNAME = 161
Private Const ERROR_INVALID_PARAMETER = &H87
Private Const ERROR_INTERNET_INVALID_URL = (INTERNET_ERROR_BASE + 5)
Private Const ERROR_INSUFFICIENT_BUFFER = 122

Private Const ERROR_NO_MORE_FILES = 18

Private Declare Function apiFormatMsgLong Lib "kernel32" _
    Alias "FormatMessageA" _
    (ByVal dwflags As Long, _
    ByVal lpSource As Long, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    Arguments As Long) _
    As Long

Private Declare Function apiFtpOpenFile Lib "wininet.dll" _
  Alias "FtpOpenFileA" _
  (ByVal hFtpSession As Long, ByVal lpszFileName As String, _
  ByVal fdwAccess As Long, ByVal dwflags As Long, _
  ByVal dwContext As Long) As Long

Private Declare Function apiFtpCreateDir Lib "wininet" _
  Alias "FtpCreateDirectoryA" _
  (ByVal hFtpSession As Long, ByVal lpszDirectory As String) _
  As Long

Private Declare Function apiFtpSetCurrentDir Lib "wininet" _
  Alias "FtpSetCurrentDirectoryA" _
  (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Long

Private Declare Function apiFTPPutFile Lib "wininet.dll" _
  Alias "FtpPutFileA" _
  (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
  ByVal lpszNewRemoteFile As String, ByVal dwflags As Long, _
  ByVal dwContext As Long) As Long

Private Declare Function apiInetCrackUrl Lib "wininet.dll" _
    Alias "InternetCrackUrlA" _
    (ByVal lpszUrl As String, ByVal dwUrlLength As Long, _
    ByVal dwflags As Long, _
    lpUrlComponents As URL_COMPONENTS) As Long

Private Declare Function apiInetCanonicalizeUrl Lib "wininet.dll" _
  Alias "InternetCanonicalizeUrlA" _
  (ByVal lpszUrl As String, ByVal lpszBuffer As String, _
  lpdwBufferLength As Long, ByVal dwflags As Long) As Long

Private Declare Function apiInetGetConnectedState Lib "wininet.dll" _
  Alias "InternetGetConnectedState" _
  (ByVal lpdwFlags As Long, ByVal dwReserved As Long) As Long

Private Declare Function apiInetReadFile Lib "wininet.dll" _
  Alias "InternetReadFile" _
  (ByVal hFile As Long, lpBuffer As Any, _
  ByVal dwBytesToRead As Long, lpBytesRead As Long) As Long

Private Declare Function apiInetReadFileStr Lib "wininet.dll" _
  Alias "InternetReadFile" _
  (ByVal hFile As Long, ByVal lpBuffer As String, _
  ByVal dwBytesToRead As Long, lpBytesRead As Long) As Long

Private Declare Function apiInetOpenURL Lib "wininet.dll" _
  Alias "InternetOpenUrlA" _
  (ByVal hInternet As Long, ByVal lpszUrl As String, _
  ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, _
  ByVal dwflags As Long, ByVal dwContext As Long) As Long

Private Declare Function apiInetOpen Lib "wininet.dll" _
  Alias "InternetOpenA" _
  (ByVal lpszAgent As String, ByVal dwAccessType As Long, _
  ByVal lpszProxy As String, ByVal lpszProxyBypass As String, _
  ByVal dwflags As Long) As Long

Private Declare Function apiInetCloseHandle Lib "wininet.dll" _
  Alias "InternetCloseHandle" _
  (ByVal hInet As Long) As Long

Private Declare Function apiInetQueryDataAvailable Lib "wininet.dll" _
  Alias "InternetQueryDataAvailable" _
  (ByVal hFile As Long, lpdwNUMBEROfBytesAvailable As Long, _
  ByVal dwflags As Long, ByVal dwContext As Long) As Long

Private Declare Function apiInetGetLastResponse Lib "wininet.dll" _
  Alias "InternetGetLastResponseInfoA" _
  (lpdwError As Long, ByVal lpszBuffer As String, _
  lpdwBufferLength As Long) As Long

Private Declare Function apiInetConnect Lib "wininet.dll" _
  Alias "InternetConnectA" _
  (ByVal hInternet As Long, ByVal lpszServerName As String, _
  ByVal nServerPort As Integer, ByVal lpszUserName As String, _
  ByVal lpszPassword As String, ByVal dwService As Long, _
  ByVal dwflags As Long, ByVal dwContext As Long) As Long

Private Declare Function apiInetAutodial Lib "wininet.dll" _
  Alias "InternetAutodial" _
  (ByVal dwflags As Long, ByVal dwReserved As Long) _
  As Long

Private Declare Function apiCreateFile Lib "kernel32.dll" _
  Alias "CreateFileA" _
  (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
  ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, _
  ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
  ByVal hTemplateFile As Long) As Long

Private Declare Function apiWriteFile Lib "kernel32.dll" _
  Alias "WriteFile" _
  (ByVal hFile As Long, lpBuffer As Any, _
  ByVal nNUMBEROfBytesToWrite As Long, _
  lpNUMBEROfBytesWritten As Long, ByVal lpOverlapped As Long) _
  As Long

Private Declare Function apiReadFile Lib "kernel32" _
  Alias "ReadFile" (ByVal hFile As Long, _
  lpBuffer As Any, ByVal nNUMBEROfBytesToRead As Long, _
  lpNUMBEROfBytesRead As Long, ByVal lpOverlapped As Long) As Long

Private Declare Function apiFlushFileBuffers Lib "kernel32.dll" _
  Alias "FlushFileBuffers" _
  (ByVal hFile As Long) As Long

Private Declare Function apiCloseHandle Lib "kernel32.dll" _
  Alias "CloseHandle" (ByVal hObject As Long) As Long

Private Declare Function apiGetActiveWindow Lib "user32" _
  Alias "GetActiveWindow" () As Long

Public Property Let AutoCreateRemoteDir(blnYesNo As Boolean)
  mblnCreateRemoteDir = blnYesNo
End Property

Private Sub sChangeDir(ByVal strDir As String)
Dim intCount As Integer
Dim strExtract As String
Dim i As Integer
Dim lngRet As Long
Const conERR_COULD_NOT_CHANGE_DIR = vbObjectError + 9999
Const conERR_NO_SUCH_DIR = 12003

  intCount = fCountWords(strDir, "/")
  If intCount > 0 Then
    For i = 2 To intCount - 1 'Last pos must be file, first one always /
      strExtract = fGetWord(strDir, i, "/")
      If strExtract = vbNullString Then
        strExtract = "/"
      End If
      lngRet = apiFtpSetCurrentDir(hSession, strExtract)
      If lngRet = 0 Then
        If Err.LastDllError = conERR_NO_SUCH_DIR Then
          If mblnCreateRemoteDir Then
            Call SysCmd(acSysCmdSetStatus, "Creating remote folder '" & strExtract & "'..")
            'Create the directory automatically
            lngRet = apiFtpCreateDir(hSession, strExtract)
            lngRet = apiFtpSetCurrentDir(hSession, strExtract)
          Else
            Err.Raise conERR_COULD_NOT_CHANGE_DIR, _
             "FTP", fInetError(Err.LastDllError)
          End If
        End If
      End If
    Next
  End If
End Sub

Public Sub UploadFileToFTPServer()
On Error GoTo ErrHandler
Dim lngRet As Long
Dim abytData() As Long
Dim lngBytesWritten As Long
Dim lngBytesRead As Long
Dim lngTotalBytesWritten As Long
Const conERR_COULD_NOT_TRANSFER_FILE = vbObjectError + 2215

  Call SysCmd(acSysCmdSetStatus, "Changing directories....")

  Call sChangeDir(mstrDestination)

  Call SysCmd(acSysCmdSetStatus, "Uploading file. Please wait...")

  lngRet = apiFTPPutFile(hSession, mstrSrcFile, mstrDestination, _
     INTERNET_FLAG_TRANSFER_BINARY Or INTERNET_FLAG_NO_CACHE_WRITE, 0&)

  If lngRet = 0 Then Err.Raise conERR_COULD_NOT_TRANSFER_FILE

  Call SysCmd(acSysCmdSetStatus, "Done...")

ExitHere:
                  On Error Resume Next
  Call SysCmd(acSysCmdClearStatus)
  Exit Sub
ErrHandler:
  Select Case Err.NUMBER
    Case conERR_COULD_NOT_TRANSFER_FILE:
      Err.Raise conERR_COULD_NOT_TRANSFER_FILE, "FTP::UploadFileToFTPServer", _
        fInetError(Err.LastDllError)
    Case Else
      With Err
        .Raise .NUMBER, .Source, fInetError(.LastDllError)
      End With
  End Select
  Resume ExitHere
End Sub

Public Sub ConnectToFTPHost( _
          Optional strUserName As String = "anonymous", _
          Optional strPassword As String = "homer@doh.com")

On Error GoTo ErrHandler
Dim lngRet As Long
Dim lngLen As Long
Dim strStatus As String
Dim lngFlags As Long
Dim strTmp As String
Dim lngFileMode As Long

  If Not fParseURL(mtURLInfo, mstrURL) = False Then
    If mtURLInfo.lpszHostName = vbNullString Then
      Err.Raise mconERR_BAD_URL
    End If
  Else
    Err.Raise mconERR_UNKNOWN
  End If

    'Begin of modification ========================
  'If the analysis of the URL contains an username and the arguments passed are the defaults
  'then use the supplied user and password in the URL instead of the default ones.

  If (mtURLInfo.lpszUserName <> "" And strUserName = "anonymous" And strPassword = "homer@doh.com") Then
    strUserName = mtURLInfo.lpszUserName
    strPassword = mtURLInfo.lpszPassword
  End If

  'End of modification ========================


  If mblnUseProxy Then
    lngFlags = INTERNET_OPEN_TYPE_PROXY
  Else
    lngFlags = INTERNET_OPEN_TYPE_PRECONFIG
  End If

  hInet = apiInetOpen("WinInet-FileTransferObjects", _
            lngFlags, _
            vbNullChar, vbNullChar, 0&)

  Call SysCmd(acSysCmdSetStatus, "Connecting to FTP Server '" & mtURLInfo.lpszHostName & "'...")

  'Begin of modification ========================
  'Use the port NUMBER obtained in the analysis of the URL instead of the default INTERNET_INVALID_PORT_NUMBER

  ' This is the original code
  'hSession = apiInetConnect(hInet, mtURLInfo.lpszHostName, _
              INTERNET_INVALID_PORT_NUMBER, _
              CStr(strUserName), _
              CStr(strPassword), INTERNET_SERVICE_FTP, _
              INTERNET_FLAG_PASSIVE, 0&)

  ' This is the modified code
  hSession = apiInetConnect(hInet, mtURLInfo.lpszHostName, _
              mtURLInfo.nPort, _
              CStr(strUserName), _
              CStr(strPassword), INTERNET_SERVICE_FTP, _
              INTERNET_FLAG_PASSIVE, 0&)

  'End of modification ========================

  If hSession = 0 Then Err.Raise mconERR_CONNECTION_FAIL

  If mblnUpload Then
    lngFileMode = GENERIC_WRITE
    Call SysCmd(acSysCmdSetStatus, "Requesting upload information from server...")
  Else
    Call SysCmd(acSysCmdSetStatus, "Requesting download information from server...")
    lngFileMode = GENERIC_READ
  End If

  If Not mblnUpload Then
    With mtURLInfo
      hFTP = apiFtpOpenFile(hSession, .lpszUrlPath, _
          lngFileMode, FTP_TRANSFER_TYPE_BINARY Or _
          INTERNET_FLAG_DONT_CACHE, 0&)
      If hFTP = 0 Then Err.Raise mconERR_CANNOT_START_TRANSFER

        lngRet = apiInetQueryDataAvailable(hFTP, mlngSize, 0&, 0&)
        If mlngSize = 0 Then mlngSize = 1
    End With
  End If

ExitHere:
                  On Error Resume Next
  Call SysCmd(acSysCmdClearStatus)
  Exit Sub
ErrHandler:
  Select Case Err.NUMBER
    Case mconERR_CANNOT_START_TRANSFER:
      Err.Raise mconERR_CANNOT_START_TRANSFER, "FTP::ConnectToFTPHost", _
        "Cannot start file transfer from '" & mtURLInfo.lpszHostName & "'."
    Case mconERR_CONNECTION_FAIL:
      Err.Raise mconERR_CONNECTION_FAIL, "FTP::ConnectToFTPHost", _
        "Couldn't connect to server '" & mtURLInfo.lpszHostName & "'."
    Case mconERR_BAD_URL:
      Err.Raise mconERR_BAD_URL, "FTP::ConnectToFTPHost", _
        "Bad URL. " & vbCrLf & mstrURL & vbCrLf & _
        "Please verify the hyperlink and try again."
    Case mconERR_UNKNOWN:
      Err.Raise mconERR_UNKNOWN, "FTP::ConnectToFTPHost", _
        "An unknown error occurred."
    Case Else:
      Err.Raise Err.NUMBER, Err.Source, Err.Description
  End Select
  Resume ExitHere
End Sub


Public Sub WriteFTPDataToFile()
On Error GoTo ErrHandler

Dim lngRet As Long
Dim lngBytesRead As Long
Dim lngFlags As Long
Dim lngBytesWritten As Long
Dim lngTotalBytesWritten As Long
Dim abytData() As Byte
Dim strHost As String
Dim strFile As String
Dim strDir As String
Const conERR_GENERIC = vbObjectError + 100
Const conFILE_EXISTS = vbObjectError + 200

  If mblnUpload Then Err.Raise mconERR_WRONG_OPERATION

  If (Me.FileExists And Not mblnOverWrite) Or mblnPromptForFile Then
    If lnghWnd = 0 Then lnghWnd = apiGetActiveWindow()
    mstrDestination = fGetNewFileName("Please select a" _
                    & " new name for the destination file.", lnghWnd, False)
    If mstrDestination = vbNullString Then Err.Raise _
      conERR_GENERIC
  End If

  Call SysCmd(acSysCmdInitMeter, "Downloading file '" & mtURLInfo.lpszUrlPath & "'...", 255)


  'Create the destination file
  hFile = apiCreateFile(mstrDestination, _
            GENERIC_READ Or GENERIC_WRITE, _
            0&, 0&, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0&)
  If hFile = INVALID_HANDLE_VALUE Then Err.Raise conERR_GENERIC

  'Read in MAX_CHUNK Chunk
  Do
    ReDim abytData(MAX_CHUNK)
    lngRet = apiInetReadFile(hFTP, _
              abytData(0), _
              MAX_CHUNK, _
              lngBytesRead)
    Call apiWriteFile(hFile, abytData(0), MAX_CHUNK, _
              lngBytesWritten, 0&)
    lngTotalBytesWritten = lngTotalBytesWritten + lngBytesWritten
    Call SysCmd(acSysCmdUpdateMeter, CInt(lngTotalBytesWritten / mlngSize))
  Loop Until lngRet <> 0 And lngBytesRead = 0

ExitHere:
                  On Error Resume Next
  Call SysCmd(acSysCmdRemoveMeter)
  Exit Sub
ErrHandler:
  Select Case Err.NUMBER
    Case conERR_GENERIC:
      'Do Nothing
    Case mconERR_WRONG_OPERATION:
      Err.Raise mconERR_WRONG_OPERATION, "FTP::WriteFTPDataToFile", _
        "Wrong transfer method selected."
    Case Else:
      Err.Raise Err.NUMBER, "FTP::WriteFTPDataToFile", Err.Description
  End Select
  Resume ExitHere
End Sub

Public Property Let SourceFile(strFileName As String)
  'Pass vbNullString for Common Dialog
  If strFileName = vbNullString Then
    If lnghWnd = 0 Then lnghWnd = apiGetActiveWindow()
    mstrSrcFile = fGetNewFileName("Please select the" _
                    & " file to upload.", lnghWnd, True)
  Else
    mstrSrcFile = strFileName
  End If
  mblnUpload = True
  mlngSize = FileLen(mstrSrcFile)
End Property

Private Sub Class_Initialize()
Dim lngFlags As Long
  mintErrorTrap = Application.GetOption("Error Trapping")
  Call Application.SetOption("Error Trapping", 2)
  mblnConnectState = CBool(apiInetGetConnectedState(lngFlags, 0&))
  Set mcolRemoteDir = New Collection
End Sub

Private Sub Class_Terminate()
                  On Error Resume Next
  Call Application.SetOption("Error Trapping", mintErrorTrap)
  Call SysCmd(acSysCmdClearStatus)
  Set mcolRemoteDir = Nothing
  Call apiInetCloseHandle(hSession)
  Call apiInetCloseHandle(hInet)
  Call apiFlushFileBuffers(hFile)
  Call apiCloseHandle(hFile)
  mlngSize = 0
End Sub


Public Property Let UseProxy(blnYesNo As Boolean)
  mblnUseProxy = blnYesNo
End Property

Public Property Let FtpURL(ByVal strURL As String)
  mstrURL = strURL
End Property

Public Property Let hWnd(ByVal SourcehWnd As Long)
  lnghWnd = SourcehWnd
End Property

Public Property Get FileExists() As Boolean
  FileExists = Not (Dir(mstrDestination) = vbNullString)
End Property

Public Property Let OverwriteTarget(ByVal blnOverWrite As Boolean)
  mblnOverWrite = blnOverWrite
End Property

Public Property Let DestinationFile(ByVal strFilePath As String)
  mstrDestination = strFilePath
End Property

Public Property Get SpecifiedURL() As String
  SpecifiedURL = mstrURL
End Property

Public Property Let PromptWithCommonDialog(blnYesNo As Boolean)
  mblnPromptForFile = blnYesNo
End Property

Public Sub About()
  DoCmd.OpenForm "wz_frm_About", windowmode:=acDialog
End Sub

Public Property Get SizeOfFile() As Long
  SizeOfFile = mlngSize
End Property

Public Property Get IsConnected() As Boolean
  IsConnected = mblnConnectState
End Property


Public Sub DialDefaultNUMBER()
Dim lngFlags As Long
Dim lngRet As Long
  lngRet = apiInetGetConnectedState(lngFlags, 0&)
  If lngRet = 0 Then
    If (lngFlags And INTERNET_CONNECTION_MODEM) Then
      'Try a connect
      lngRet = apiInetAutodial(INTERNET_AUTODIAL_FORCE_UNATTENDED, 0&)
      mblnConnectState = Not (lngRet = 0)
    End If
  End If
End Sub


Private Function fCountWords(ByVal s As String, ByVal strDelim As String) As Integer
' Counts the words in a string that are separated by commas.
'Modified from MS KB
'
Dim WC As Integer, Pos As Integer
  WC = 1
  Pos = InStr(s, strDelim)
  Do While Pos > 0
    WC = WC + 1
    Pos = InStr(Pos + 1, s, strDelim)
  Loop
  fCountWords = WC
End Function

Private Function fGetWord(ByVal s As String, Indx As Integer, strDelim As String)
' Returns the nth word in a specific field.
'Modified from MS KB
'
Dim WC As Integer, COUNT As Integer, SPos As Integer, EPos As Integer
  WC = fCountWords(s, strDelim)
  If Indx < 1 Or Indx > WC Then
    fGetWord = vbNullString
    Exit Function
  End If
  COUNT = 1
  SPos = 1
  For COUNT = 2 To Indx
    SPos = InStr(SPos, s, strDelim) + 1
  Next COUNT
  EPos = InStr(SPos, s, strDelim) - 1
  If EPos <= 0 Then EPos = Len(s)
  fGetWord = Trim(MID(s, SPos, EPos - SPos + 1))
End Function

Private Function fParseURL(tURLInfo As URL_COMPONENTS, ByVal strURL As String) As Boolean
Dim lngLen As Long, strBuffer As String
Dim lngRet As Long, strURLLocal As String

  strBuffer = String$(MAX_BUFFER, 0)
  lngLen = Len(strBuffer)
  lngRet = apiInetCanonicalizeUrl(strURL, strBuffer, lngLen, ICU_BROWSER_MODE)
  If Not lngRet = 0 Then
    strURLLocal = Left$(strBuffer, lngLen)
    With tURLInfo
      .lpszScheme = String$(MAX_BUFFER, 0)
      .dwSchemeLength = MAX_BUFFER
      .nScheme = INTERNET_SCHEME_UNKNOWN
      .lpszHostName = String$(MAX_BUFFER, 0)
      .dwHostNameLength = MAX_BUFFER
      .dwStructSize = Len(tURLInfo)
      .nPort = 0
      .lpszUserName = String$(MAX_BUFFER, 0)
      .dwUserNameLength = MAX_BUFFER
      .lpszPassword = String$(MAX_BUFFER, 0)
      .dwPasswordLength = MAX_BUFFER
      .lpszUrlPath = String$(MAX_BUFFER, 0)
      .dwUrlPathLength = MAX_BUFFER
      .lpszExtraInfo = String$(MAX_BUFFER, 0)
      .dwExtraInfoLength = MAX_BUFFER
    End With
    lngRet = apiInetCrackUrl(strURLLocal, Len(strURLLocal), _
                      ICU_ESCAPE, tURLInfo)
    If lngRet = 0 Then
      fParseURL = False
    Else
      fParseURL = True
      With tURLInfo
        .lpszExtraInfo = fTrimNull(.lpszExtraInfo)
        .lpszHostName = fTrimNull(.lpszHostName)
        .lpszPassword = fTrimNull(.lpszPassword)
        .lpszScheme = fTrimNull(.lpszScheme)
        .lpszUrlPath = fTrimNull(.lpszUrlPath)
        .lpszUserName = fTrimNull(.lpszUserName)
      End With
    End If
  End If
End Function

Private Function fInetError(ByVal lngErrCode As Long) As String
Dim lngLen As Long, strBuffer As String
  Call apiInetGetLastResponse(lngErrCode, vbNullString, lngLen)
  strBuffer = String$(lngLen + 1, 0)
  Call apiInetGetLastResponse(lngErrCode, strBuffer, lngLen)
  fInetError = strBuffer
End Function

Private Function fTrimNull(ByVal strIn As String) As String
  fTrimNull = Left$(strIn, InStr(strIn, vbNullChar) - 1)
End Function

Private Function fGetNewFileName(strMsg As String, hWnd As Long, blnOpen As Boolean) As String
Dim clsDialog As cDialog
  Set clsDialog = New cDialog
  With clsDialog
    .TITLE = strMsg
    .hWnd = hWnd
    .ModeOpen = blnOpen
    .StartDir = CurDir()
    fGetNewFileName = .Action
  End With
  Set clsDialog = Nothing
End Function

Private Function fAPIErr(ByVal lngErr As Long) As String
'Original Idea obtained from
'Hardcode Visual Basic 5
'by Bruce McKinney
'
Dim strMsg As String
Dim lngRet As Long
    strMsg = String$(1024, 0)
    lngRet = apiFormatMsgLong(FORMAT_MESSAGE_FROM_SYSTEM, 0&, _
                           lngErr, 0&, strMsg, Len(strMsg), ByVal 0&)
    If Not lngRet = 0 Then
        fAPIErr = Left$(strMsg, lngRet)
    End If
End Function
