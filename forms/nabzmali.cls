Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(CANCEL As Integer)
                On Error Resume Next
       Forms![BASEKNOW]![hhwin] = Me.hWnd
    DoCmd.RunSQL "create  view nabz_mali7 as SELECT top(7) DEED_HED.DATE_S, SUM(DEED_DTL.BED) AS mab FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE (DEED_DTL.HES = N'" & Forms![BASEKNOW]![hesnaghd] & "') GROUP BY DEED_HED.DATE_S ORDER BY DEED_HED.DATE_S DESC"
    DoCmd.RunSQL "alter   view nabz_mali7 as SELECT top(7) DEED_HED.DATE_S, SUM(DEED_DTL.BED) AS mab FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE (DEED_DTL.HES = N'" & Forms![BASEKNOW]![hesnaghd] & "') GROUP BY DEED_HED.DATE_S ORDER BY DEED_HED.DATE_S DESC"
    DoCmd.RunSQL "create  view nabz_mali30 as SELECT top(30) DEED_HED.DATE_S, SUM(DEED_DTL.BED) AS mab, dbo.Uday(dbo.DEED_HED.DATE_S) AS dd, dbo.Umonth(dbo.DEED_HED.DATE_S) AS mm FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE (DEED_DTL.HES = N'" & Forms![BASEKNOW]![hesnaghd] & "') GROUP BY DEED_HED.DATE_S , dbo.Uday(dbo.DEED_HED.DATE_S), dbo.Umonth(dbo.DEED_HED.DATE_S) ORDER BY DEED_HED.DATE_S DESC"
    DoCmd.RunSQL "alter   view nabz_mali30 as SELECT top(30) DEED_HED.DATE_S, SUM(DEED_DTL.BED) AS mab, dbo.Uday(dbo.DEED_HED.DATE_S) AS dd, dbo.Umonth(dbo.DEED_HED.DATE_S)  AS mm FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE (DEED_DTL.HES = N'" & Forms![BASEKNOW]![hesnaghd] & "') GROUP BY DEED_HED.DATE_S, dbo.Uday(dbo.DEED_HED.DATE_S), dbo.Umonth(dbo.DEED_HED.DATE_S) ORDER BY DEED_HED.DATE_S DESC"
    DoCmd.RunSQL "create  view nabz_mali90 as SELECT top(90) DEED_HED.DATE_S, SUM(DEED_DTL.BED) AS mab, dbo.Uday(dbo.DEED_HED.DATE_S) AS dd, dbo.Umonth(dbo.DEED_HED.DATE_S) AS mm FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE (DEED_DTL.HES = N'" & Forms![BASEKNOW]![hesnaghd] & "') GROUP BY DEED_HED.DATE_S, dbo.Uday(dbo.DEED_HED.DATE_S), dbo.Umonth(dbo.DEED_HED.DATE_S) ORDER BY DEED_HED.DATE_S DESC"
    DoCmd.RunSQL "alter   view nabz_mali90 as SELECT top(90) DEED_HED.DATE_S, SUM(DEED_DTL.BED) AS mab, dbo.Uday(dbo.DEED_HED.DATE_S) AS dd, dbo.Umonth(dbo.DEED_HED.DATE_S)  AS mm FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE (DEED_DTL.HES = N'" & Forms![BASEKNOW]![hesnaghd] & "') GROUP BY DEED_HED.DATE_S, dbo.Uday(dbo.DEED_HED.DATE_S) , dbo.Umonth(dbo.DEED_HED.DATE_S) ORDER BY DEED_HED.DATE_S DESC"
'----------------------
  '  DoCmd.RunSQL "create  view nabz_asnad7 as SELECT top(7) DEED_HED.DATE_S, SUM(DEED_DTL.BED) AS mab FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE (DEED_DTL.HES = N'" & Forms![baseknow]![ADA] & "') GROUP BY DEED_HED.DATE_S"
  '  DoCmd.RunSQL "alter  view nabz_asnad7 as SELECT top(7) DEED_HED.DATE_S, SUM(DEED_DTL.BED) AS mab FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE (DEED_DTL.HES = N'" & Forms![baseknow]![ADA] & "') GROUP BY DEED_HED.DATE_S"
  '
  '  DoCmd.RunSQL "create  view nabz_asnad30 as SELECT top(30) DEED_HED.DATE_S, SUM(DEED_DTL.BED) AS mab, dbo.Uday(dbo.DEED_HED.DATE_S) AS dd, dbo.Umonth(dbo.DEED_HED.DATE_S) AS mm FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE (DEED_DTL.HES = N'" & Forms![baseknow]![ADA] & "') GROUP BY DEED_HED.DATE_S, dbo.Uday(dbo.DEED_HED.DATE_S) , dbo.Umonth(dbo.DEED_HED.DATE_S)"
  '  DoCmd.RunSQL "alter  view nabz_asnad30 as SELECT top(30) DEED_HED.DATE_S, SUM(DEED_DTL.BED) AS mab, dbo.Uday(dbo.DEED_HED.DATE_S) AS dd, dbo.Umonth(dbo.DEED_HED.DATE_S) AS mm FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE (DEED_DTL.HES = N'" & Forms![baseknow]![ADA] & "') GROUP BY DEED_HED.DATE_S, dbo.Uday(dbo.DEED_HED.DATE_S) , dbo.Umonth(dbo.DEED_HED.DATE_S)"
'
''    DoCmd.RunSQL "create  view nabz_asnad90 as SELECT top(90) DEED_HED.DATE_S, SUM(DEED_DTL.BED) AS mab, dbo.Uday(dbo.DEED_HED.DATE_S) AS dd, dbo.Umonth(dbo.DEED_HED.DATE_S) AS mm FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE (DEED_DTL.HES = N'" & Forms![baseknow]![ADA] & "') GROUP BY DEED_HED.DATE_S, dbo.Uday(dbo.DEED_HED.DATE_S) , dbo.Umonth(dbo.DEED_HED.DATE_S)"
 '   DoCmd.RunSQL "alter  view nabz_asnad90 as SELECT top(90) DEED_HED.DATE_S, SUM(DEED_DTL.BED) AS mab, dbo.Uday(dbo.DEED_HED.DATE_S) AS dd, dbo.Umonth(dbo.DEED_HED.DATE_S) AS mm FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S WHERE (DEED_DTL.HES = N'" & Forms![baseknow]![ADA] & "') GROUP BY DEED_HED.DATE_S, dbo.Uday(dbo.DEED_HED.DATE_S) , dbo.Umonth(dbo.DEED_HED.DATE_S)"

End Sub
