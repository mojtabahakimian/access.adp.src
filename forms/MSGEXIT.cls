Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit
Dim LETEXIT As Boolean


Private Sub Command1_Click()
   LETEXIT = True
   DoCmd.Quit
End Sub

Private Sub Command17_Click()
     LETEXIT = False
    DoCmd.OpenForm Me.NAME, , , , , acHidden
End Sub

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
   LETEXIT = False

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub


Private Sub Form_Timer()
               On Error Resume Next ' IDLEMINUTES determines how much idle time to wait for before
         ' running the IdleTimeDetected subroutine.
         Dim IDLEMINUTES As Integer
  If MID(Forms![BASEKNOW]![OPTIONSS], 39, 1) = 5 Then
         IDLEMINUTES = Val(MID(Forms![BASEKNOW]![OPTIONSS], 40, 2))
         If IDLEMINUTES = 0 Then
            IDLEMINUTES = 10
         End If
         Static PrevControlName As String
         Static PrevFormName As String
         Static ExpiredTime
         Dim ActiveFormName As String
         Dim ActiveControlName As String
         Dim ExpiredMinutes

                         On Error Resume Next

         ' Get the active form and control name.

         ActiveFormName = Screen.ActiveForm.NAME
         If Err Then
            ActiveFormName = "No Active Form"
            Err = 0
         End If

         ActiveControlName = Screen.ActiveControl.NAME
            If Err Then
            ActiveControlName = "No Active Control"
            Err = 0
         End If

         ' Record the current active names and reset ExpiredTime if:
         '    1. They have not been recorded yet (code is running
         '       for the first time).
         '    2. The previous names are different than the current ones
         '       (the user has done something different during the timer
         '        interval).
         If (PrevControlName = "") Or (PrevFormName = "") Or (ActiveFormName <> PrevFormName) Or (ActiveControlName <> PrevControlName) Then
            PrevControlName = ActiveControlName
            PrevFormName = ActiveFormName
            ExpiredTime = 0
         Else
            ' ...otherwise the user was idle during the time interval, so
            ' increment the total expired time.
            ExpiredTime = ExpiredTime + Me.TimerInterval
         End If

         ' Does the total expired time exceed the IDLEMINUTES?
         ExpiredMinutes = (ExpiredTime / 1000) / 60
         If ExpiredMinutes >= IDLEMINUTES Then
            ' ...if so, then reset the expired time to zero...
            ExpiredTime = 0
            ' ...and call the IdleTimeDetected subroutine.
            Err.Clear
            ActiveFormName = Screen.ActiveForm.NAME
            If Err = 0 Then
                While Not ActiveFormName = "No Active Form" And Err = 0
                    DoCmd.Close acForm, ActiveFormName, acSaveYes
                    ActiveFormName = Screen.ActiveForm.NAME
                Wend
            End If
            Err.Clear
            ActiveFormName = Screen.ActiveReport.NAME
            If Err = 0 Then
                While Not ActiveFormName = "No Active Form" And Err = 0
                    DoCmd.Close acReport, ActiveFormName, acSaveYes
                    ActiveFormName = Screen.ActiveReport.NAME
                Wend
            End If
            Err.Clear
            DoCmd.OpenForm "user_login"
         End If
    End If
End Sub

Private Sub Form_Unload(CANCEL As Integer)
                On Error Resume Next ' IDLEMINUTES determines how much idle time to wait for before
  DoCmd.OpenForm Me.NAME, , , , , acDialog
  If LETEXIT = True Then
    CANCEL = False
  Else
    CANCEL = True
  End If
End Sub
