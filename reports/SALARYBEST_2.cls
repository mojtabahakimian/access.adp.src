Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
  '  Dim rst As New ADODB.Recordset
  '  rst.Open "SELECT  N_S FROM dbo.WORKHEAD WHERE (dbo.Umonth(WDATE) = " & Forms![F_MENU_BIM]![MMO] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  '  If rst.RecordCount > 0 Then
  '      Me.RecordSource = "SELECT TOP 100 PERCENT n_s,HES_K, HES_M, HES_T, DBO.FISHSTR(ISNULL(SHARH,' ')) AS SHARH, BED, BES FROM DEED_DTL WHERE (N_S = " & rst.Fields("N_S") & ") and (HES_K = " & Forms![Baseknow]![PERSONEL] & ") AND (HES_M = 1) AND (BES > 0) ORDER BY HES_M"
  '  End If
    'Me.RecordSource = "SELECT     TOP 100 PERCENT DEED_DTL.HES_K, DEED_DTL.HES_M, DEED_DTL.HES_T, DEED_DTL.SHARH, DEED_DTL.BED, DEED_DTL.BES, dbo.Umonth(DEED_HED.DATE_S) AS MM, DEED_HED.DATE_S FROM         DEED_HED INNER JOIN  DEED_DTL ON DEED_HED.N_S = DEED_DTL.N_S WHERE (DEED_DTL.HES_K = " & Forms![Baseknow]![PERSONEL] & ") And (DEED_DTL.HES_M = 1) And (dbo.UMonth(DEED_HED.DATE_S) = 5) And (DEED_HED.NO_S = 11) And (DEED_DTL.BES > 0) ORDER BY DEED_DTL.HES_M, DEED_HED.DATE_S"
End Sub
