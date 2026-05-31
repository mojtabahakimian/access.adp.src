Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Private Sub Report_Error(DataErr As Integer, Response As Integer)
            On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
Dim rst As New ADODB.Recordset
         On Error Resume Next
        Me.RecordSource = "SELECT INVO_LST.CODE, SUM(INVO_LST.MEGH) AS SMEGH, SUM(INVO_LST.MEGHk) AS SMEGHK, TCOD_VAHEDS.NAMES AS vahed, STUF_DEF.NAME + N' ' + ISNULL(INVO_LST.MANDAH, N'') AS KALA, INVO_LST.MANDAH, SUM(INVO_LST.MABL_K) AS smabk, TOZIE.TID, TOZIE.TDATE, TOZIE.TDRIVER, TOZIE.TCITY, TOZIE.TMAMUR, INVO_LST.ANBAR FROM TOZIE_SUB INNER JOIN TOZIE ON TOZIE_SUB.TID = TOZIE.TID INNER JOIN INVO_LST INNER JOIN STUF_DEF ON INVO_LST.CODE = STUF_DEF.CODE INNER JOIN TCOD_VAHEDS ON INVO_LST.VAHED_K = TCOD_VAHEDS.CODE ON TOZIE_SUB.NUMBER = INVO_LST.NUMBER WHERE (INVO_LST.TAG = 2) GROUP BY INVO_LST.CODE, TCOD_VAHEDS.NAMES, STUF_DEF.NAME + N' ' + ISNULL(INVO_LST.MANDAH, N''), INVO_LST.MANDAH, TOZIE.TID, TOZIE.TDATE, TOZIE.TDRIVER, TOZIE.TCITY, TOZIE.TMAMUR, INVO_LST.ANBAR  HAVING (TOZIE.TID = " & Me.OpenArgs & ")"

        rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight  FROM         dbo.INVO_LST INNER JOIN                      dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN                      dbo.TOZIE_SUB ON dbo.INVO_LST.NUMBER = dbo.TOZIE_SUB.NUMBER WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.TOZIE_SUB.TID = " & Me.OpenArgs & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            If Not IsNull(rst.Fields("Weight")) Then
                Me.VAZN.CAPTION = "وزن كل به كيلو : " & Round(rst.Fields("Weight"))
           End If
       End If
   DoCmd.Maximize
End Sub
