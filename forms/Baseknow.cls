Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                 On Error Resume Next
     If DataErr = 17 Or DataErr = 11 Or DataErr = 4060 Then
          Forms![BASEKNOW]![Text44] = False
          DoCmd.OpenForm "msgdialog", acNormal, , , acFormReadOnly, acDialog, "دستري به سرور مقدور نمي باشد . شبكه قطع است يا سرور خاموش است . اتصالات شبكه و كابلها را كنترل كنيد. يا نام سرور صحيح نيست درصوتيكه مايليد نام سرور را وارد كنيد "
          If Forms![BASEKNOW]![Text44] Then
            DoCmd.RunCommand acCmdConnection
          Else
             DoCmd.Quit
          End If
     Else
                          On Error Resume Next: errdetector (DataErr)
     End If
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub


Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next
   DoCmd.OpenForm "user_login", , , , , acDialog
    Me.dt = FARSIDATE2(DATE)
    Me.dttmp = Me.dt
    Me.Requery
End Sub

Private Sub Form_Timer()
                 On Error Resume Next
    Me.Form.TimerInterval = 500000
    Me.dt = FARSIDATE2(DATE)
    Me.dttmp = Me.dt
    Me.Requery
End Sub
