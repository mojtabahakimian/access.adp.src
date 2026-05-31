Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim CTL As Controls

Private Sub dt1_AfterUpdate()
    Me.tozieselect.Form.RecordSource = "SELECT HEAD_LST.NUMBER, HEAD_LST.DATE_N, CUST_HESAB.NAME, CUST_HESAB.hes, CUST_HESAB.ADDRESS FROM CUST_HESAB INNER JOIN HEAD_LST ON CUST_HESAB.hes = HEAD_LST.CUST_NO WHERE (HEAD_LST.DATE_N BETWEEN " & DT1 & " AND " & DT2 & ") AND (HEAD_LST.TAG = 2) ORDER BY HEAD_LST.NUMBER"

End Sub

Private Sub dt2_AfterUpdate()
    Me.tozieselect.Form.RecordSource = "SELECT HEAD_LST.NUMBER, HEAD_LST.DATE_N, CUST_HESAB.NAME, CUST_HESAB.hes, CUST_HESAB.ADDRESS FROM CUST_HESAB INNER JOIN HEAD_LST ON CUST_HESAB.hes = HEAD_LST.CUST_NO WHERE (HEAD_LST.DATE_N BETWEEN " & DT1 & " AND " & DT2 & ") AND (HEAD_LST.TAG = 2) ORDER BY HEAD_LST.NUMBER"

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    DoCmd.Close
  End If
End Sub
