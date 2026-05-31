Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Command13_Click()
                    On Error Resume Next
  Dim rst As New ADODB.Recordset
  rst.Open "SQLSTATE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  rst.AddNew
  rst.Fields("SQLST") = Me.OpenArgs
  rst.Fields("TITEL") = Me.TITEL
  rst.Fields("USER_NAME") = UCurrentUser
  rst.Fields("LETOTHER") = IIf(Me.C1, 1, 0)
  rst.Fields("CR_DATE") = Now()
  rst.update
  rst.Close
  DoCmd.OpenForm "MESAG", , , , , acDialog, "گزارش با موفقيت ذخيره شد"
  DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command2_Click()
                       On Error Resume Next
    DoCmd.Close
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
