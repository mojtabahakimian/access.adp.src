Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
                     On Error Resume Next
        If Me.mab > 0 Then
         COMM.CAPTION = "چكهاي دريافت شده " & Me.CountOfMABL & " فقره جمعاًبه مبلغ :" & Format(Me.mab, "#,### ريال") & "  "
        End If
        Err.Clear
End Sub

Private Sub Report_Close()
                     On Error Resume Next
  DoCmd.Close acForm, "F_MENU_KHFR"
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Report_Open(CANCEL As Integer)
                     On Error Resume Next
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
