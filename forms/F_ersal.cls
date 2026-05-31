Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command2_Click()
                      On Error Resume Next
   Dim errorno As Integer, dbname, tablename, sql As String
   dbname = Text7
   If (RIGHT(dbname, 1) <> "\") Then
      dbname = dbname & "\"
   End If
   DoCmd.Close acForm, "baseknow"
   DoCmd.Close acForm, "ye"
   DoCmd.Close acForm, "DEFAULT"

   DoCmd.OpenForm "baseknow", acNormal, , , acFormEdit, acHidden
    If (MID(Forms![BASEKNOW]![tindata], 20, 7) <> "CORRECT" Or IsNull(Forms![BASEKNOW]![tindata])) And False Then
      DoCmd.OpenForm "YE"
    End If
   DoCmd.OpenForm "DEFAULT"
End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "F_ERSAL", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "F_ERSAL", 3)
End Sub
