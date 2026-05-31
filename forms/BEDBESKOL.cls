Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Load()
   DoCmd.Maximize

End Sub

Private Sub Form_Open(CANCEL As Integer)
   Me.RecordSource = "SELECT     TOP 100 PERCENT HES_K, SUM(SumOfBED) AS SumOfBED, SUM(SumOfBES) AS SumOfBES, dbo.TOTA_HES.NAME, dbo.UIIF(SUM(BEDBES), '>', 0, SUM(BEDBES), 0) AS BED,dbo.UIIF(SUM(BEDBES), '>', 0, 0, SUM(BEDBES) * - 1) AS BES FROM         dbo.BEDBESMAH" & Forms![BASEKNOW]![USERCOD] & " INNER JOIN    dbo.TOTA_HES ON HES_K = dbo.TOTA_HES.NUMBER GROUP BY HES_K, dbo.TOTA_HES.NAME ORDER BY HES_K"
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub شماره_حساب_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "BEDEHBESTANMOIN", acNormal
End Sub
