Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
 Me.RecordSource = "SELECT SUM(MEGHKF) AS MEGHF, SUM(MABK) AS MABKK, SUM(MEGHkK) AS MEGHk, CUST_NO, MPNAME FROM FROOSH_COUNTALL_CITY_N_FANI(" & Forms![F_MENU_DATE]![DT1] & "," & Forms![F_MENU_DATE]![DT2] & ") FROOSH_COUNTALL_CITY_N_FANI GROUP BY CUST_NO, MPNAME ORDER BY CUST_NO"
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
