Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub PageFooter_Format(CANCEL As Integer, FormatCount As Integer)
  If Page = Pages Then
    Me.Label231.Visible = False
    Me.Label232.Visible = True
  End If
End Sub

Private Sub Report_Close()
                     On Error Resume Next
   DoCmd.Close acForm, "F_MENU_BIM"
End Sub


Private Sub Report_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                    On Error Resume Next
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub ReportFooter_Format(CANCEL As Integer, FormatCount As Integer)
  If Page = Pages Then
    Me.Label231.Visible = False
    Me.Label232.Visible = True
  End If
End Sub
