Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Close()
                    On Error Resume Next
   DoCmd.Close acForm, "F_MENU_CHEK"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
    Me.RecordSource = "SELECT     PAY_GETD.*, TCOD_BANKS.NAMES  FROM         PAY_GETD INNER JOIN     TCOD_BANKS ON PAY_GETD.BANK = TCOD_BANKS.CODE WHERE     (PAY_GETD.N_KOL <> 911) AND (PAY_GETD.N_KOL <> " & Forms![BASEKNOW]![BANKHA] & ") OR                      (NOT (PAY_GETD.N_KOL2 IS NULL)) OR     (NOT (PAY_GETD.N_KOL3 IS NULL)) ORDER BY PAY_GETD.N_SERI "
End Sub
