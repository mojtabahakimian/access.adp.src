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
  Me.RecordSource = "SELECT SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHKF) AS MEGHF, SUM(FROOSH_COUNTALL_CITY_N_FANI.MABK) AS MABKK, SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHkK) AS MEGHk, TCOD_OSTAN.OSNAME, FROOSH_COUNTALL_CITY_N_FANI.OSCO FROM         FROOSH_COUNTALL_CITY_N_FANI(" & Forms![F_MENU_DATE]![DT1] & "," & Forms![F_MENU_DATE]![DT2] & ") FROOSH_COUNTALL_CITY_N_FANI INNER JOIN TCOD_OSTAN ON FROOSH_COUNTALL_CITY_N_FANI.OSCO = TCOD_OSTAN.OSCODE GROUP BY TCOD_OSTAN.OSNAME, FROOSH_COUNTALL_CITY_N_FANI.OSCO ORDER BY SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHKF)"
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
Private Sub MEGHF_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "FROOSH_COUNTALL_CITY1_OSTAN_CUSTO", acFormDS
End Sub
Private Sub OSNAME_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "FROOSH_COUNTALL_CITY1_OSTAN_CUSTO", acFormDS, , , , , Me.OSCO
End Sub
