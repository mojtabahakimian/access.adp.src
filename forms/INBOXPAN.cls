Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Type FLASHWINFO
    cbSize As Long
    hWnd As Long
    dwflags As Long
    uCount As Long
    dwTimeout As Long
End Type
Private Declare Function FindWindowA Lib "user32.dll" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FlashWindowEx Lib "user32.dll" (pfwi As FLASHWINFO) As Long
Private Declare Function FlashWindow Lib "user32.dll" (ByVal hWnd As Long, bInvert As Long) As Long
Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal uFlag As Long) As Long
Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

Option Compare Database
Private Sub Command28_Click()
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT * FROM MESAGEP WHERE idnum = " & Me.IDNUM, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
            rst.Fields("STATUS") = 2
            rst.update
    End If
    DoCmd.Close acForm, Me.NAME
End Sub


Private Sub Command33_Click()
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT * FROM MESAGEP WHERE idnum = " & Me.IDNUM, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
            rst.Fields("STATUS") = 2
            rst.update
    End If
    Me.TimerInterval = 0
    DoCmd.OpenForm "mesagepanel", , , , , acDialog, 1
End Sub

Private Sub Command37_Click()
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT * FROM MESAGEP WHERE idnum = " & Me.IDNUM, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
            rst.Fields("STATUS") = 2
            rst.update
    End If
    Me.TimerInterval = 0
    DoCmd.OpenForm "mesagepanel", , , , , acDialog, 3
End Sub

Private Sub Form_Close()
    End_Tray Me

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Open(CANCEL As Integer)
    Dim rst As New ADODB.Recordset, VS As String
    VS = ""
    rst.Open "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL=0) and idd <> 120", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        VS = VS & rst.Fields("IDD") & ";" & Chr(34) & DECODEUN(rst.Fields("sal_name")) & Chr(34)
        rst.MoveNext
        If Not rst.EOF Then
            VS = VS & ";"
        End If
    Wend
    Me.PERSONEL.RowSource = VS
    Add_Tray Me, "" & Chr(0)
    ShowNotifyBalloonTip IIf(IsNull(Me.PAYAM), "پيام خالي", Me.PAYAM), "نرم افزار شفيق : پيام جديد از " & Me.USERNAME, 4, 30000
End Sub


Private Sub Form_Timer()
   Dim RETVAL As Boolean, fEnable As Long, acc As Long
    If Me.PAYAM = "كاربر محترم سلام نرم افزار شما جهت آبديت  تا  10 ثانيه ديگر بصورت خودكار بسته خواهد شد بعد از چند دقيقه ميتوانيد دوباره وارد شويد با تشكر" Then
       Me.IDNUM.SetFocus
       Me.Command28.Enabled = False
       Me.Command33.Enabled = False
       Me.Command37.Enabled = False
       Me.TimerInterval = 0
       Dim rst As New ADODB.Recordset
       rst.Open "SELECT * FROM MESAGEP WHERE idnum = " & Me.IDNUM, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 Then
            rst.Fields("STATUS") = 2
            rst.update
       End If
       t = Timer
       Do
            Me.CAPTION = 10 - Int(Timer - t)
            DoEvents
       Loop Until Timer > t + 10
       DoCmd.OpenForm "MSGEXIT", acNormal, , , , , 0
       DoCmd.Quit
   End If

    fEnable = 5000
  ' RETVAL = FlashWindow(Me.hWnd, fEnable)
  ' RETVAL = FlashWindow(Application.hWndAccessApp, fEnable)
 ' RETVAL = SetActiveWindow(Application.hWndAccessApp)
 '  RETVAL = EnableWindow(Application.hWndAccessApp, 1)
   DoCmd.RunCommand acCmdAppMaximize
End Sub

Private Sub PAYAM_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "zoom", acFormDS, , , , , "1" & Me.PAYAM
End Sub
