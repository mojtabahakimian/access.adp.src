Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Report_Close()
    DoCmd.Close acForm, "F_MENU_ANBAR_FRKH_GRP"
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
   On Error Resume Next
   Me.RecordSource = "SELECT     dbo.Q_LIST_FCT2.ADDRESS,dbo.Q_LIST_FCT2.NUMBER, dbo.Q_LIST_FCT2.NAME, dbo.Q_LIST_FCT2.SumOfMABL_K, dbo.Q_LIST_FCT2.DATE_N, dbo.Q_LIST_FCT2.TAG,dbo.Q_LIST_FCT2.AGHLAM, dbo.Q_LIST_FCT2.hes, ISNULL(dbo.[BED-BES].BEST, 0) AS mand, dbo.CUST_HESAB.CITY, dbo.TOZIE.TDATE,dbo.TOZIE.TDRIVER , dbo.TOZIE.TCITY, dbo.TOZIE.TID, dbo.TOZIE.TMAMUR FROM dbo.Q_LIST_FCT2 INNER JOIN dbo.CUST_HESAB ON dbo.Q_LIST_FCT2.hes = dbo.CUST_HESAB.hes INNER JOIN dbo.TOZIE_SUB ON dbo.Q_LIST_FCT2.NUMBER = dbo.TOZIE_SUB.NUMBER INNER JOIN dbo.TOZIE ON dbo.TOZIE_SUB.TID = dbo.TOZIE.TID LEFT OUTER JOIN dbo.[BED-BES] ON dbo.Q_LIST_FCT2.hes = dbo.[BED-BES].HES WHERE     (dbo.TOZIE.TID = " & Me.OpenArgs & " AND (Q_LIST_FCT2.TAG = 2)) "
   DoCmd.Maximize
End Sub
