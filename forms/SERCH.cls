Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Dim nam As String

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
   If KeyAscii = 27 Then
      DoCmd.Close acForm, Me.NAME
   End If
   If KeyAscii = 13 Then
      If IsNull(Text0) Then
            DoCmd.OpenForm "SERCHBED", acFormDS, , , , , nam
      Else
            DoCmd.OpenForm "SERCHBED", acFormDS, , "NAME LIKE " & "'%" & Text0 & "%'", , , nam
      End If
      DoCmd.Close acForm, Me.NAME
   End If
End Sub

Private Sub Form_Load()
                      On Error Resume Next
   nam = Screen.Application.CurrentObjectName
   DoCmd.Beep
End Sub
