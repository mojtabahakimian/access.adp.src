Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Command3_Click()
                   On Error Resume Next
  DoCmd.Close acForm, "BASEKNOW", acSaveYes
  ChangeDB ("NEGIN" & Me.YEA)
  DoCmd.OpenForm "BASEKNOW", , , , , acHidden
  DoCmd.Close acForm, "YE"
  DoCmd.OpenForm "YE"
  DoCmd.Close acForm, Me.NAME
End Sub
Private Sub Command4_Click()
                    On Error Resume Next
 DoCmd.Close
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "SELECTYEAR", 3)
End Sub
Private Sub YEA_AfterUpdate()
                  On Error Resume Next
 If Me.YEA = "+" Or Me.YEA = "-" Then
   DoCmd.OpenForm "SELECTYEAR", acNormal, , , , acDialog
 End If
End Sub
