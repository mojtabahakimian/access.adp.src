Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Report_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
'  ' DEBUG.PRINT "SELECT     SUM(ABS(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES)) AS sumofbed, dbo.DEED_DTL.HES_K, dbo.Umonth(dbo.DEED_HED.DATE_S) AS mm FROM         dbo.DEED_DTL INNER JOIN    dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S GROUP BY dbo.DEED_DTL.HES_K, dbo.Umonth(dbo.DEED_HED.DATE_S) HAVING      (dbo.DEED_DTL.HES_K = " & Forms![F_MENU_KOL_DATE]![HKOL] & ") "
'  Me.OLEUnbound0.RowSource = "SELECT     SUM(ABS(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES)) AS sumofbed, dbo.DEED_DTL.HES_K, dbo.Umonth(dbo.DEED_HED.DATE_S) AS mm FROM         dbo.DEED_DTL INNER JOIN    dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S GROUP BY dbo.DEED_DTL.HES_K, dbo.Umonth(dbo.DEED_HED.DATE_S) HAVING      (dbo.DEED_DTL.HES_K = " & Forms![F_MENU_KOL_DATE]![HKOL] & ") "
End Sub
