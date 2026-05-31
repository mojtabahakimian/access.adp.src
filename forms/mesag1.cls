Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Sub Command0_Click()
On Error GoTo Err_Command0_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command0_Click:
    Exit Sub

Err_Command0_Click:
    MsgBox Err.Description
    Resume Exit_Command0_Click

End Sub

Private Sub Command1_Click()
                  On Error Resume Next
  DoCmd.Close acForm, "mesag"
End Sub


Private Sub Form_Load()
                  On Error Resume Next
  DoCmd.restore
  If Left(Me.OpenArgs, 99) = "اين فاكتور خريد بااين شماره قبلا ثبت شده است دقت كنيد كه فاكتور تكراري نباشد : شماره فاكتور قبلي : " Then
    DoCmd.Beep
    Me.Detail.BackColor = 255
    Me.MSG.ForeColor = 65535
  End If
  If Left(Me.OpenArgs, 36) = "قيمت فروش از قيمت خريد كمتر مي باشد." Then
    DoCmd.Beep
    Me.Detail.BackColor = 255
    Me.MSG.ForeColor = 65535
  End If
  If Left(Me.OpenArgs, 58) = "اخطار مهم درصد محاسبه سهم طرف ديگر مشخص نشده است براي مشخص" Then
    DoCmd.Beep
    Me.Detail.BackColor = 255
    Me.MSG.ForeColor = 16777215
  End If

  If Left(Me.OpenArgs, 54) = "سقف اعتبار مشتري تمام شده است اعتبار اين مشتري مبلغ : " Then
    DoCmd.Beep
    Me.Detail.BackColor = 255
    Me.MSG.ForeColor = 16777215
  End If

End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
    MSG = "xcvbgc"
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
