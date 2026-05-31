Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim nam As String
Private Sub Form_KeyPress(KeyAscii As Integer)
                   On Error Resume Next
   If KeyAscii = 27 Then
      DoCmd.Close acForm, Me.NAME
   End If
   If KeyAscii = 13 Then
      If IsNull(Text0) Then
            DoCmd.OpenForm "JIS_TASK", acFormDS
      Else
            Select Case Me.OpenArgs
                Case 1:
                    DoCmd.OpenForm "JIS_TASK", acFormDS, , "IDNUM =" & Text0
                Case 2:
                    DoCmd.OpenForm "JIS_TASK", acFormDS, , "TASK LIKE N'%" & Text0 & "%'"
                Case 3:
                    DoCmd.OpenForm "JIS_TASK2", acFormDS, , , , , Text0
            End Select
      End If
      DoCmd.Close acForm, Me.NAME
   End If
End Sub

Private Sub Form_Load()
                  On Error Resume Next
   If IsNull(Me.OpenArgs) Then
        nam = Screen.Application.CurrentObjectName
   Else
     nam = Me.OpenArgs
   End If
   DoCmd.Beep
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
