Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command5_Click()
                      On Error Resume Next
   If IsNull(Me.DT1) And IsNull(Me.DT2) Then
       DoCmd.OpenForm "CHEK_SERIAL_CONTROL", acFormDS
   Else
      If IsNull(Me.DT1) And Not IsNull(Me.DT2) Then
           DoCmd.OpenForm "CHEK_SERIAL_CONTROL", acFormDS, , "N_SERI <= " & Me.DT2
      Else
           If IsNull(Me.DT2) And Not IsNull(Me.DT1) Then
                DoCmd.OpenForm "CHEK_SERIAL_CONTROL", acFormDS, , "N_SERI >=  " & Me.DT1
           Else
                DoCmd.OpenForm "CHEK_SERIAL_CONTROL", acFormDS, , "N_SERI >=  " & Me.DT1 & " AND N_SERI <= " & Me.DT2
           End If
      End If
   End If
   DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command6_Click()
                       On Error Resume Next
    DoCmd.Close
End Sub

Private Sub Form_Load()
                     On Error Resume Next
  Select Case Me.OpenArgs
    Case "SUDZIAN": Me.HelpContextId = 301
    Case "ZIANHADE": Me.HelpContextId = 302
  End Select
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
