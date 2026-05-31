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
    Me.RecordSource = "SELECT     PAY_GETp.*, TCOD_BANKS.NAMES  FROM         PAY_GETp INNER JOIN     TCOD_BANKS ON PAY_GETp.BANK = TCOD_BANKS.CODE WHERE        (PAY_GETP.N_KOL <> 911) AND (NOT (PAY_GETP.N_KOL2 IS NULL)) OR  (PAY_GETP.N_KOL <> 911) AND (NOT (PAY_GETP.N_KOL3 IS NULL))"
End Sub
