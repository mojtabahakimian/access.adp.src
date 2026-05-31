Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit


Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                      On Error Resume Next
     If [Forms]![head_lst_pishfroosh]![TAKHFIF] > 0 Then
       Me.Label155.CAPTION = Round([Forms]![head_lst_pishfroosh]![TAKHFIF] / [Forms]![head_lst_pishfroosh]![Text59] * 100) & "% تخفيف :"
     End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
