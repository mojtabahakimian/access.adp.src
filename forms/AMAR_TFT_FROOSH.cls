Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(CANCEL As Integer)
            On Error Resume Next
  ' DEBUG.PRINT "SELECT     SUM(dbo.DEED_DTL.BED) AS mabl, dbo.DEPART.DEPATMAN, dbo.DEPART.DEPNAME FROM         dbo.DEED_DTL INNER JOIN dbo.HEAD_LST ON dbo.DEED_DTL.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.DEED_DTL.TAG = dbo.HEAD_LST.TAG INNER JOIN   dbo.DEPART ON dbo.HEAD_LST.DEPATMAN = dbo.DEPART.DEPATMAN WHERE     (dbo.HEAD_LST.DATE_N BETWEEN " & Forms![F_MENU_DATE]![DT1] & " AND " & Forms![F_MENU_DATE]![DT2] & ") AND (dbo.DEED_DTL.HES_K = " & Forms![baseknow]![TFROSH] & ") AND (dbo.DEED_DTL.TAG = 2)GROUP BY dbo.DEPART.DEPATMAN, dbo.DEPART.DEPNAME"
  Me.RecordSource = "SELECT     SUM(dbo.DEED_DTL.BED) AS mabl, dbo.DEPART.DEPATMAN, dbo.DEPART.DEPNAME FROM         dbo.DEED_DTL INNER JOIN dbo.HEAD_LST ON dbo.DEED_DTL.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.DEED_DTL.TAG = dbo.HEAD_LST.TAG INNER JOIN   dbo.DEPART ON dbo.HEAD_LST.DEPATMAN = dbo.DEPART.DEPATMAN WHERE     (dbo.HEAD_LST.DATE_N BETWEEN " & Forms![F_MENU_DATE]![DT1] & " AND " & Forms![F_MENU_DATE]![DT2] & ") AND (dbo.DEED_DTL.HES_K = " & Forms![BASEKNOW]![TFROSH] & ") AND (dbo.DEED_DTL.TAG = 2)GROUP BY dbo.DEPART.DEPATMAN, dbo.DEPART.DEPNAME"
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
