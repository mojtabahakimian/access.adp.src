Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub dt1_AfterUpdate()
   Me.RecordSource = "SELECT     TOP 100 PERCENT dbo.CUST_HESAB.NAME, dbo.STUF_DEF.NAME AS kala, dbo.INVO_LST.TAG, dbo.TAGCOD.BARGAH, dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_NO, SUM(dbo.INVO_LST.MEGH_MAR) AS mar, SUM(dbo.INVO_LST.MEGHk) AS mrgh FROM dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN  dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN  dbo.TAGCOD ON dbo.INVO_LST.TAG = dbo.TAGCOD.CODE WHERE     (dbo.HEAD_LST.DATE_N >= " & Me.DT1 & ") AND (dbo.HEAD_LST.DATE_N <= " & Me.DT2 & ") AND (dbo.INVO_LST.CODE = '378' OR  dbo.INVO_LST.CODE =  '375')  GROUP BY dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_NO, dbo.CUST_HESAB.NAME, dbo.INVO_LST.TAG, dbo.STUF_DEF.NAME, dbo.TAGCOD.BARGAH ORDER BY dbo.CUST_HESAB.NAME, dbo.STUF_DEF.NAME, dbo.INVO_LST.TAG"
End Sub

Private Sub dt2_AfterUpdate()
   Me.RecordSource = "SELECT     TOP 100 PERCENT dbo.CUST_HESAB.NAME, dbo.STUF_DEF.NAME AS kala, dbo.INVO_LST.TAG, dbo.TAGCOD.BARGAH, dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_NO, SUM(dbo.INVO_LST.MEGH_MAR) AS mar, SUM(dbo.INVO_LST.MEGHk) AS mrgh FROM dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN  dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN  dbo.TAGCOD ON dbo.INVO_LST.TAG = dbo.TAGCOD.CODE WHERE     (dbo.HEAD_LST.DATE_N >= " & Me.DT1 & ") AND (dbo.HEAD_LST.DATE_N <= " & Me.DT2 & ") AND (dbo.INVO_LST.CODE = '378' OR  dbo.INVO_LST.CODE =  '375' )  GROUP BY dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_NO, dbo.CUST_HESAB.NAME, dbo.INVO_LST.TAG, dbo.STUF_DEF.NAME, dbo.TAGCOD.BARGAH ORDER BY dbo.CUST_HESAB.NAME, dbo.STUF_DEF.NAME, dbo.INVO_LST.TAG"
End Sub

Private Sub Form_Open(CANCEL As Integer)
Dim rst As New ADODB.Recordset
  rst.Open "SELECT     HEAD_LST_1.TAG, dbo.HEAD_LST.TAG AS Expr1, dbo.HEAD_LST.NUMBER, HEAD_LST_1.NUMBER AS Expr2, HEAD_LST_1.CUST_NO,   dbo.HEAD_LST.CUST_NO AS Expr3 FROM         dbo.HEAD_LST INNER JOIN                       dbo.HEAD_LST HEAD_LST_1 ON dbo.HEAD_LST.NUMBER = HEAD_LST_1.NUMBER AND dbo.HEAD_LST.CUST_NO <> HEAD_LST_1.CUST_NO WHERE     (HEAD_LST_1.TAG = 12) AND (dbo.HEAD_LST.TAG = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
     DoCmd.OpenForm "mesag", , , , , acDialog, "فاكتور خريد با رسيد ها در نام مشتري مغايرت دارد لطفا اول مغايرت را اصلاح كنيد"
     DoCmd.OpenForm "R_F_MO", acFormDS
  End If
End Sub
