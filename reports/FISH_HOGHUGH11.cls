Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Dim JAM As Long
Option Explicit

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


'Public Function PRINT_MORA(dt As Long, PCOD As Long)
'       Dim MORA As Long
'       MORA = MORAKH_MAND(dt, PCOD)
'       If MORA <> 110110110 Then
'            PRINT_MORA = ROUND([MORA] / 440) & " _ " & ROUND((([MORA] / 440) - ROUND([MORA] / 440)) * 440 / 60) & " : " & ROUND((((([MORA] / 440) - ROUND([MORA] / 440)) * 440 / 60) - ROUND((([MORA] / 440) - ROUND([MORA] / 440)) * 440 / 60)) / 60)
'        Else
'            PRINT_MORA = 0
'       End If
'End Function
