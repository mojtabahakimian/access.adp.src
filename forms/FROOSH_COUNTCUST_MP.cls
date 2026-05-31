Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CUST_NO_DblClick(CANCEL As Integer)
      DoCmd.OpenReport "list_froosh_cust_size_tarh", acPreview, , "CUST_NO = '" & Me.CUST_NO & "'"
End Sub

Private Sub Form_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "F_MENU_ANBAR"
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      DoCmd.OpenForm "FROOSH_COUNTCUST_MP1", acFormDS, , "CUST_NO = '" & Me.CUST_NO & "'"
   End If
  If KeyAscii = 27 Then
     DoCmd.Close acForm, Me.NAME
  End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
 ' Me.RecordSource = "SELECT SUM(MEGHKF) AS MEGHF, SUM(MABK) AS MABKK, SUM(MEGHkK) AS MEGHk, CUST_NO, CUST_HESAB.NAME  FROM FROOSH_COUNTALL_CITY_N_FANI(" & Forms![F_MENU_DATE]![DT1] & "," & Forms![F_MENU_DATE]![DT2] & ") FROOSH_COUNTALL_CITY_N_FANI GROUP BY CUST_NO, CUST_HESAB.NAME"
 Me.RecordSource = "SELECT     SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHKF) AS MEGHF, SUM(FROOSH_COUNTALL_CITY_N_FANI.MABK) AS MABKK,SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHkK) AS MEGHk, FROOSH_COUNTALL_CITY_N_FANI.CUST_NO, CUST_HESAB.NAME FROM         FROOSH_COUNTALL_CITY_N_FANI(" & Forms![F_MENU_DATE]![DT1] & "," & Forms![F_MENU_DATE]![DT2] & ") FROOSH_COUNTALL_CITY_N_FANI INNER JOIN CUST_HESAB ON FROOSH_COUNTALL_CITY_N_FANI.CUST_NO = CUST_HESAB.hes GROUP BY FROOSH_COUNTALL_CITY_N_FANI.CUST_NO, CUST_HESAB.NAME ORDER BY SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHkK)DESC"
  End Sub
Private Sub SAIZE_DblClick(CANCEL As Integer)
      DoCmd.OpenForm "FROOSH_COUNTCUST_MP1", acFormDS, , "CUST_NO = " & Me.CUST_NO
End Sub
Private Sub Text35_DblClick(CANCEL As Integer)
      DoCmd.OpenReport "list_froosh_cust_size_tarh", acPreview
End Sub
