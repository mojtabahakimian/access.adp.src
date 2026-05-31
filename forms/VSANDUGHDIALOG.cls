Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Private Sub Command1_Click()
                    On Error Resume Next
 If IsNull(Me.VAZ) Or Me.VAZ = "" Then
  DoCmd.OpenForm "mesag", , , , , acDialog, "صندوقي كه چك بايد به آن منتقل شود مشخص نگرديده است...!"
 Else
  Forms![BASEKNOW]![Text44] = True
  DoCmd.OpenForm Me.NAME, , , , , acHidden
 End If
End Sub
Private Sub Command17_Click()
                    On Error Resume Next
   Forms![BASEKNOW]![Text44] = False
   DoCmd.Close acForm, Me.NAME
End Sub


Private Sub Form_Load()
                     On Error Resume Next
  Me.VAZ.RowSource = "SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = " & GETKOL(Forms![BASEKNOW]![ADA]) & ") AND (NUMBER = " & GETMOIN(Forms![BASEKNOW]![ADA]) & ")"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub
