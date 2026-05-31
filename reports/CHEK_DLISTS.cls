Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Close()
                     On Error Resume Next
    DoCmd.Close acForm, "F_MENU_CHEK"
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                 On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                     On Error Resume Next
    Me.RecordSource = "SELECT     PAY_GETD.N_SERI, PAY_GETD.BANK, PAY_GETD.DATE_S, PAY_GETD.DATE, PAY_GETD.SHOBEH, PAY_GETD.MABL, PAY_GETD.NAME_TAH,  PAY_GETD.N_HESAB, PAY_GETD.N_S, TCOD_BANKS.NAMES, PAY_GETD.RADIF, PAY_GETD.N_KOL, PAY_GETD.N_MOIN, PAY_GETD.N_KOL2, PAY_GETD.N_MOIN2, PAY_GETD.N_KOL3, PAY_GETD.N_MOIN3, PAY_GETD.N_TAF, PAY_GETD.N_TAF2, PAY_GETD.N_TAF3, TDETA_HES.NAME FROM TCOD_BANKS INNER JOIN  PAY_GETD ON TCOD_BANKS.CODE = PAY_GETD.BANK LEFT OUTER JOIN TDETA_HES ON PAY_GETD.N_KOL = TDETA_HES.N_KOL AND PAY_GETD.N_MOIN = TDETA_HES.NUMBER AND   PAY_GETD.N_TAF = TDETA_HES.TNUMBER WHERE     (PAY_GETD.N_KOL = " & Forms![BASEKNOW]![BANKHA] & " OR  PAY_GETD.N_KOL IS NULL) AND (PAY_GETD.N_KOL2 IS NULL) AND (PAY_GETD.N_KOL3 IS NULL) AND  DATE_S >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE_S <= " & Forms![F_MENU_CHEK]![DT2] & "  AND NAME_TAH LIKE '%" & Forms![F_MENU_CHEK]![MMOIN] & "%'"
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
