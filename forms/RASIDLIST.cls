Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
                      On Error Resume Next
   DoCmd.RunCommand (acCmdSelectRecord)
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
  If KeyAscii = 13 Then
     Forms![head_lst_khareed1]![NUMBER] = Me.NUMBER
     Forms![head_lst_khareed1]![CUST_NO] = Me.CUST_NO
     Forms![head_lst_khareed1]![FNUMCO] = Me.FNUMCO
  End If
  DoCmd.Close acForm, Me.NAME
End Sub
