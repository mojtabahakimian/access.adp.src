Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub COMP_COD_DblClick(CANCEL As Integer)
                    On Error Resume Next
    DoCmd.OpenReport "TASKS", acViewPreview, , "COMP_COD = '" & Me.COMP_COD & "'"
End Sub

Private Sub COMP_COD_NotInList(NewData As String, Response As Integer)
                    On Error Resume Next
  Me.COMP_COD = NewData
  Response = acDataErrContinue
End Sub

Private Sub Form_Current()
                 On Error Resume Next
    If Me.skid <> 0 Then
        LockedALL (True)
    Else
        LockedALL (False)
    End If
    If Not IsNull(Me.IDNUM) Then
       Me.AllowEdits = False
       Me.EVENTs.Locked = True
       Me.EVENTs.Form.AllowEdits = False
    Else
       Me.AllowEdits = True
       Me.EVENTs.Locked = False
       Me.EVENTs.Form.AllowEdits = True
     End If
    If Me.PERSONEL = [Forms]![BASEKNOW]![USERCOD] And Me.SEE = 0 And Not Me.NewRecord Then
        Me.SEE = 1
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
Dim rst As New ADODB.Recordset
                    On Error Resume Next
    If Not IsNull(Me.OpenArgs) Then
        If ISHESAB3(Me.OpenArgs) Then
            Me.RecordSource = "SELECT TASKS.* FROM TASKS WHERE (((TASKS.COMP_COD)='" & Me.OpenArgs & "')) ORDER BY IDNUM"
            Me.COMP_COD.DefaultValue = Me.OpenArgs
        Else
            Me.RecordSource = "SELECT TASKS.* FROM TASKS WHERE (((TASKS.username)=N'" & Me.OpenArgs & "')) ORDER BY IDNUM"
        End If
      Me.COMP_COD.Locked = True
      DoCmd.GoToRecord acDataForm, Me.NAME, acLast
    End If
    rst.Open "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        VS = VS & rst.Fields("IDD") & ";" & Chr(34) & DECODEUN(rst.Fields("sal_name")) & Chr(34)
        rst.MoveNext
        If Not rst.EOF Then
            VS = VS & ";"
        End If
    Wend
    Me.PERSONEL.RowSource = VS
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub IDNUM_DblClick(CANCEL As Integer)
    Dim RETVAL
                    On Error Resume Next

Dim fs, B
Dim pathm As String
'pathm = Forms![baseknow]![WIDTH_D] & Me.IDNUM
'   Set fs = CreateObject("Scripting.FileSystemObject")
 '  Set b = fs.CreateFolder(pathm)
If UCase(Left(Forms![BASEKNOW]![WIDTH_D], 3)) = "FTP" Then
  Dim objFTP As ftp
  Set objFTP = New ftp
  With objFTP
    .FtpURL = Forms![BASEKNOW]![WIDTH_D]
    .ConnectToFTPHost "ali", "1354", Me.IDNUM
  End With
  RETVAL = Shell("explorer.exe /e," & Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, 1)
Else
    If (Dir(Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, vbDirectory) <> "") Then
        RETVAL = Shell("explorer.exe /e," & Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, 1)
    Else
        MkDir (Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM)
        RETVAL = Shell("explorer.exe /e," & Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, 1)
   End If
End If

End Sub

Private Sub PERSONEL_DblClick(CANCEL As Integer)
    Dim dt As Date
                On Error Resume Next
    If Not IsNull(Me.IDNUM) Then
        dt = Now()
        If Forms![BASEKNOW]![TRANSF] Then
            Call TR("TASKS", "(IDNUM = " & Me.IDNUM & " )", dt, 1)
            Call TR("EVENTS", "(IDNUM = " & Me.IDNUM & " )", dt, 1)
       End If
       Me.AllowEdits = True
       Me.EVENTs.Enabled = True
       Me.EVENTs.Locked = False
       Me.EVENTs.Form.AllowEdits = True
       Me.EVENTs.Form.AllowAdditions = True
       Me.EVENTs.Form.Refresh
    End If

End Sub

Private Sub STATUS_AfterUpdate()
  Me.ENDATE = FARSIDATE(DATE)
  Me.ENTIME = Hour(Now()) * 100 + Minute(Now())
End Sub
Private Sub LockedALL(fl As Boolean)
                 On Error Resume Next
'    Me.PERSONEL.Locked = fl
    Me.TASK.Locked = fl
    Me.COMP_COD.Locked = fl
End Sub
