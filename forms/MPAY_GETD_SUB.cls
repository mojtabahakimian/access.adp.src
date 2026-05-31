Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_Current()
                   On Error Resume Next
  Me.N_KOL.RowSource = "SELECT     NUMBER, NAME FROM TOTA_HES WHERE (NUMBER = " & Forms![BASEKNOW]![BANKHA] & ")ORDER BY NAME"
  Me.N_KOL.Requery
  Me.N_MOIN.RowSource = "SELECT     DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE     (((DETA_HES.N_KOL) =" & Me.N_KOL & ")) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA_HES.NAME"
  Me.N_MOIN.Requery
  Me.N_TAF.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBER) =" & Me.N_MOIN & ") AND ((TDETA_HES.N_KOL) =" & Me.N_KOL & "))GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.NAME"
  Me.N_TAF.Requery
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Me.N_KOL.RowSource = "SELECT     NUMBER, NAME FROM TOTA_HES WHERE (NUMBER = " & Forms![BASEKNOW]![BANKHA] & ")ORDER BY NAME"
  Me.N_KOL.Requery
  Me.N_MOIN.RowSource = "SELECT     DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE     (((DETA_HES.N_KOL) =" & Me.N_KOL & ")) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA_HES.NAME"
  Me.N_MOIN.Requery
  Me.N_TAF.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBER) =" & Me.N_MOIN & ") AND ((TDETA_HES.N_KOL) =" & Me.N_KOL & "))GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.NAME"
  Me.N_TAF.Requery
End Sub
