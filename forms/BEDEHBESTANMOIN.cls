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

Private Sub Form_Open(CANCEL As Integer)
                     On Error Resume Next
Me.RecordSource = "SELECT     TOP 100 PERCENT HES_K, SUM(SumOfBED) AS SumOfBED, SUM(SumOfBES) AS SumOfBES,  dbo.UIIF(SUM(BEDBES), '>', 0, SUM(BEDBES), 0) AS BED, dbo.UIIF(SUM(BEDBES), '>', 0, 0, SUM(BEDBES) * - 1) AS BES, HES_M, MOIN FROM         BEDBESMAH" & Forms![BASEKNOW]![USERCOD] & " INNER JOIN   dbo.TOTA_HES ON HES_K = dbo.TOTA_HES.NUMBER WHERE     (HES_K = " & Forms![BEDBESKOL]![شماره حساب] & ") GROUP BY HES_K, HES_M, MOIN ORDER BY HES_K, HES_M"
'' DEBUG.PRINT "SELECT     TOP 100 PERCENT HES_K, SUM(SumOfBED) AS SumOfBED, SUM(SumOfBES) AS SumOfBES,  dbo.UIIF(SUM(BEDBES), '>', 0, SUM(BEDBES), 0) AS BED, dbo.UIIF(SUM(BEDBES), '>', 0, 0, SUM(BEDBES) * - 1) AS BES, HES_M, MOIN FROM         BEDBESMAH" & Forms![BASEKNOW]![USERCOD] & " INNER JOIN   dbo.TOTA_HES ON HES_K = dbo.TOTA_HES.NUMBER WHERE     (HES_K = " & Forms![BEDBESKOL]![شماره حساب] & ") GROUP BY HES_K, HES_M, MOIN ORDER BY HES_K, HES_M"

DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub معين_DblClick(CANCEL As Integer)
                     On Error Resume Next
    DoCmd.OpenForm "BEDEHKARAN_BESTANKARAN_MAHDOD", acFormDS
End Sub
