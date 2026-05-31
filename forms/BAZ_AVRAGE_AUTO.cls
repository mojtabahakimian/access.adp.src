Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Command14_Click()
                      On Error Resume Next
   DoCmd.Close
End Sub

Private Sub Form_Load()

   Forms![AUTO_RUN].Form.TimerInterval = 0
   Dim errorno As Integer, sql As String, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst As New ADODB.Recordset, rst4 As New ADODB.Recordset
   Dim CON, i As Long, MIAN As Double, MBKM As Double, MOGUDI As Double
  Dim Temp As Double
                    On Error Resume Next
   rst4.Open "SELECT TCOD_STUFGROUP.CODE, TCOD_STUFGROUP.NAMES FROM TCOD_STUFGROUP WHERE (((TCOD_STUFGROUP.CODE)<>0)) ORDER BY TCOD_STUFGROUP.NAMES", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst4.EOF
        rst3.Open "INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If IsNull(Me.co) Or Me.co = "" Then
           rst.Open "SELECT     TOP 100 PERCENT dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR, dbo.STUF_FSK.MOGODI_A, dbo.STUF_FSK.FI_A, dbo.STUF_FSK.MABL_A FROM  dbo.STUF_DEF INNER JOIN   dbo.STUF_FSK ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_DEF.RADAH =  " & rst4.Fields(0) & " ) ORDER BY dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Else
           rst.Open "SELECT     TOP 100 PERCENT dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR, dbo.STUF_FSK.MOGODI_A, dbo.STUF_FSK.FI_A, dbo.STUF_FSK.MABL_A FROM  dbo.STUF_DEF INNER JOIN   dbo.STUF_FSK ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_DEF.RADAH =  " & rst4.Fields(0) & " and dbo.STUF_FSK.CODE = '" & Me.co & "' and dbo.STUF_FSK.ANBAR = 2) ORDER BY dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        End If
        Me.Text23 = rst.RecordCount
        Select Case rst4.Fields(0)
             Case "2":
                     While Not rst.EOF
                         Me.co = rst.Fields("CODE")
                         MIAN = GETSTANDARDPRICE(rst.Fields("CODE"))
                         If MIAN <> 0 Then
                             CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS  WHERE TABLE_NAME = 'TMPAV101')   DROP VIEW TMPAV101")
                             CurrentProject.Connection.Execute ("CREATE VIEW  TMPAV101 as " & _
                               " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, dbo.INVO_LST.TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,  dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO,dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM  dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG = dbo.TAGCOD.CODE WHERE  (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH , dbo.INVO_LST.NUMBER UNION " & _
                               " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, 6 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBARF AS ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG + 1 = dbo.TAGCOD.CODE " & _
                               " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBARF = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & Me.dt & ")  AND (dbo.INVO_LST.TAG = 5) ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
                               " SELECT     TOP 100 PERCENT dbo.HEAD_LST_FBK.DATE_N, 4 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN       dbo.HEAD_LST_FBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_FBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_FBK.dtag INNER JOIN  dbo.TAGCOD ON dbo.HEAD_LST_FBK.htag = dbo.TAGCOD.CODE " & _
                               " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_FBK.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST_FBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
                               " SELECT     TOP 100 PERCENT dbo.HEAD_LST_KBK.DATE_N, 3 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A,dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE, dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH  FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST_KBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_KBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_KBK.dtag INNER JOIN " & _
                               " dbo.TAGCOD ON dbo.HEAD_LST_KBK.htag = dbo.TAGCOD.CODE WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_KBK.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST_KBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER")
                             Set RST2 = New ADODB.Recordset
                             RST2.Open "SELECT     TOP 100 PERCENT DATE_N, TAG, NUMBER, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID,  MEGH_R , RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, ID, BARGAH FROM dbo.TMPAV101 ORDER BY DATE_N, BARGAH", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                             While Not RST2.EOF
                                  Text19 = Text19 + 1
                                  Me.Repaint
                                  rst3.Filter = "ID = " & RST2.Fields("ID")
                                  Set rst3 = rst3
                                  Select Case RST2.Fields("TAG")
                                      Case 1: 'خريد
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.update
                                      Case 22: 'برگشت فروش سال قبل
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.update
                                      Case 2: 'فروش
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.update
                                      Case 5: 'انتقالي خروج
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.Fields("MABL") = MIAN
                                           rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                           rst3.update
                                      Case 6: 'انتقالي خروج
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.Fields("MABL") = MIAN
                                           rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                           rst3.update
                                      Case 10: 'مواد خروج
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.Fields("MABL") = MIAN
                                           rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                           rst3.update
                                      Case 11:    'موادساير خروج
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.Fields("MABL") = MIAN
                                           rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                           rst3.update
                                      Case 9:    'موادساير خروج
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.Fields("MABL") = MIAN
                                           rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                           rst3.update
                                  End Select
                                  RST2.MoveNext
                               Wend
                               If Me.FORMOL Then
                                   DoCmd.RunSQL "UPDATE    dbo.DTL_MANF  SET  MABLK = (MEGHk + PERT) * " & MIAN & ", SMABL = " & MIAN & " WHERE     (CODE = N'" & rst.Fields("CODE") & "')"
                               End If
                            End If
                            rst.MoveNext
                      Wend
             Case "3":
                     While Not rst.EOF
                         Me.co = rst.Fields("CODE")
                         MIAN = GETSTANDARDPRICE(rst.Fields("CODE"))
                         If MIAN <> 0 Then
                             CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS  WHERE TABLE_NAME = 'TMPAV101')   DROP VIEW TMPAV101")
                             CurrentProject.Connection.Execute ("CREATE VIEW  TMPAV101 as " & _
                               " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, dbo.INVO_LST.TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,  dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO,dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM  dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG = dbo.TAGCOD.CODE WHERE  (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH , dbo.INVO_LST.NUMBER UNION " & _
                               " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, 6 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBARF AS ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG + 1 = dbo.TAGCOD.CODE " & _
                               " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBARF = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & Me.dt & ")  AND (dbo.INVO_LST.TAG = 5) ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
                               " SELECT     TOP 100 PERCENT dbo.HEAD_LST_FBK.DATE_N, 4 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN       dbo.HEAD_LST_FBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_FBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_FBK.dtag INNER JOIN  dbo.TAGCOD ON dbo.HEAD_LST_FBK.htag = dbo.TAGCOD.CODE " & _
                               " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_FBK.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST_FBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
                               " SELECT     TOP 100 PERCENT dbo.HEAD_LST_KBK.DATE_N, 3 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A,dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE, dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH  FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST_KBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_KBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_KBK.dtag INNER JOIN " & _
                               " dbo.TAGCOD ON dbo.HEAD_LST_KBK.htag = dbo.TAGCOD.CODE WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_KBK.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST_KBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER")
                             Set RST2 = New ADODB.Recordset
                             RST2.Open "SELECT     TOP 100 PERCENT DATE_N, TAG, NUMBER, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID,  MEGH_R , RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, ID, BARGAH FROM dbo.TMPAV101 ORDER BY DATE_N, BARGAH", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                             While Not RST2.EOF
                                  Text19 = Text19 + 1
                                  Me.Repaint
                                  rst3.Filter = "ID = " & RST2.Fields("ID")
                                  Set rst3 = rst3
                                  Select Case RST2.Fields("TAG")
                                      Case 1: 'خريد
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.update
                                      Case 22: 'برگشت فروش سال قبل
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.update
                                      Case 2: 'فروش
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.update
                                      Case 6: 'انتقالي خروج
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.Fields("MABL") = MIAN
                                           rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                           rst3.update
                                      Case 5: 'انتقالي خروج
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.Fields("MABL") = MIAN
                                           rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                           rst3.update
                                      Case 10: 'مواد خروج
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.Fields("MABL") = MIAN
                                           rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                           rst3.update
                                      Case 11:    'موادساير خروج
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.Fields("MABL") = MIAN
                                           rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                           rst3.update
                                      Case 9:    'موادساير خروج
                                           rst3.Fields("AVRAGE") = MIAN
                                           rst3.Fields("MABL") = MIAN
                                           rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                           rst3.update
                                  End Select
                                  RST2.MoveNext
                               Wend
                               If Me.FORMOL Then
                                   DoCmd.RunSQL "UPDATE    dbo.DTL_MANF  SET  MABLK = (MEGHk + PERT) * " & MIAN & ", SMABL = " & MIAN & " WHERE     (CODE = N'" & rst.Fields("CODE") & "')"
                               End If
                            End If
                            rst.MoveNext
                      Wend
             Case Else:
                 While Not rst.EOF
                     Me.co = rst.Fields("CODE")
                     MBKM = rst.Fields("MABL_A")
                     MIAN = rst.Fields("FI_A")
                     MOGUDI = rst.Fields("MOGODI_A")
                     CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS  WHERE TABLE_NAME = 'TMPAV101')   DROP VIEW TMPAV101")
                     CurrentProject.Connection.Execute ("CREATE VIEW  TMPAV101 as " & _
                       " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, dbo.INVO_LST.TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,  dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO,dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM  dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG = dbo.TAGCOD.CODE WHERE  (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH , dbo.INVO_LST.NUMBER UNION " & _
                       " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, 6 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBARF AS ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG + 1  = dbo.TAGCOD.CODE " & _
                       " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBARF = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & Me.dt & ")  AND (dbo.INVO_LST.TAG = 5) ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
                       " SELECT     TOP 100 PERCENT dbo.HEAD_LST_FBK.DATE_N, 4 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN       dbo.HEAD_LST_FBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_FBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_FBK.dtag INNER JOIN  dbo.TAGCOD ON dbo.HEAD_LST_FBK.htag = dbo.TAGCOD.CODE " & _
                       " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_FBK.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST_FBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
                       " SELECT     TOP 100 PERCENT dbo.HEAD_LST_KBK.DATE_N, 3 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A,dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE, dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH  FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST_KBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_KBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_KBK.dtag INNER JOIN " & _
                       " dbo.TAGCOD ON dbo.HEAD_LST_KBK.htag = dbo.TAGCOD.CODE WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_KBK.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST_KBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER")
                     Set RST2 = New ADODB.Recordset
                     RST2.Open "SELECT     TOP 100 PERCENT DATE_N, TAG, NUMBER, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID,  MEGH_R , RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, ID, BARGAH FROM dbo.TMPAV101 ORDER BY DATE_N, BARGAH", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                     While Not RST2.EOF
                          Text19 = Text19 + 1
                          Me.Repaint
                          rst3.Filter = "ID = " & RST2.Fields("ID")
                          Set rst3 = rst3
                          Select Case RST2.Fields("TAG")
                              Case 1: 'خريد
                                   MBKM = MBKM + RST2.Fields("MABL_K")
                                   MOGUDI = MOGUDI + RST2.Fields("MEGHK")
                                   If MBKM = 0 Then
                                      MIAN = 0
                                   Else
                                      If MOGUDI = 0 Then
                                          MIAN = 0
                                          MBKM = 0
                                      Else
                                          MIAN = MBKM / MOGUDI
                                      End If
                                   End If
                                   rst3.Fields("AVRAGE") = MIAN
                                   rst3.update
                              Case 22: 'برگشت فروش سال قبل
                                   MBKM = MBKM + RST2.Fields("MABL") * RST2.Fields("MEGH_MAR")
                                   MOGUDI = MOGUDI + RST2.Fields("MEGH_MAR")
                                   If MBKM = 0 Then
                                      MIAN = 0
                                   Else
                                      If MOGUDI = 0 Then
                                          MIAN = 0
                                          MBKM = 0
                                      Else
                                          MIAN = MBKM / MOGUDI
                                      End If
                                   End If
                                   rst3.Fields("AVRAGE") = MIAN
                                   rst3.update
                              Case 2: 'فروش
                                   MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                                   MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                                   rst3.Fields("AVRAGE") = MIAN
                                   rst3.update
                              Case 3: 'برگشت خريد
                                   MBKM = MBKM - (RST2.Fields("MEGH_MAR") * MIAN)
                                   MOGUDI = MOGUDI - RST2.Fields("MEGH_MAR")
                                   If MBKM = 0 Then
                                      MIAN = 0
                                   Else
                                      If MOGUDI = 0 Then
                                          MIAN = 0
                                          MBKM = 0
                                      Else
                                          MIAN = MBKM / MOGUDI
                                      End If
                                   End If
                                   rst3.Fields("AVRAGE2") = MIAN
                                   rst3.update
                              Case 4: 'برگشت فروش
                                   MBKM = MBKM + (RST2.Fields("MEGH_MAR") * MIAN)
                                   MOGUDI = MOGUDI + RST2.Fields("MEGH_MAR")
                                   If MBKM = 0 Then
                                      MIAN = 0
                                   Else
                                      If MOGUDI = 0 Then
                                          MIAN = 0
                                          MBKM = 0
                                      Else
                                          MIAN = MBKM / MOGUDI
                                      End If
                                   End If
                                   rst3.Fields("AVRAGE2") = MIAN
                                   rst3.update
                              Case 5: 'انتقالي خروج
                                   MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                                   MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                                   rst3.Fields("AVRAGE") = MIAN
                                   rst3.Fields("MABL") = MIAN
                                   rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                   rst3.update
                              Case 6: 'انتقالي ورود
                                   MBKM = MBKM + RST2.Fields("MABL_K")
                                   MOGUDI = MOGUDI + RST2.Fields("MEGHK")
                                   If MBKM = 0 Then
                                      MIAN = 0
                                   Else
                                      If MOGUDI = 0 Then
                                          MIAN = 0
                                          MBKM = 0
                                      Else
                                          MIAN = MBKM / MOGUDI
                                      End If
                                   End If
                                   rst3.Fields("AVRAGE2") = MIAN
                                   rst3.update
                              Case 10: 'مواد خروج
                                   MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                                   MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                                   rst3.Fields("AVRAGE") = MIAN
                                   rst3.Fields("MABL") = MIAN
                                   rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                   rst3.update
                              Case 11:    'موادساير خروج
                                   MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                                   MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                                   rst3.Fields("AVRAGE") = MIAN
                                   rst3.Fields("MABL") = MIAN
                                   rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                   rst3.update
                          End Select
                          RST2.MoveNext
                       Wend
                       rst.MoveNext
                 Wend
        End Select
        rst.Close
        'rst2.Close
        rst3.Close
        rst4.MoveNext

   Wend
   Forms![AUTO_RUN].Form.TimerInterval = 100000
    DoCmd.Close acForm, Me.NAME

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Timer()
                     On Error Resume Next
   Me.Refresh
End Sub

Private Sub RADAH_AfterUpdate()
  If RADAH = 2 Or RADAH = 3 Then
     Me.FORMOL.Visible = True
     Me.Label26.Visible = True
  Else
     Me.FORMOL.Visible = False
     Me.Label26.Visible = False
  End If
End Sub
