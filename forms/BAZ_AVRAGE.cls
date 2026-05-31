Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub co_AfterUpdate()
    Dim rst As New ADODB.Recordset
    If Not IsNull(Me.co) Then
        rst.Open "select code , radah from stuf_def where code = '" & Me.co & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 1 Then
            Me.RADAH = rst.Fields("radah")
        Else
            DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "كد يافت نشد"
        End If
    End If
End Sub

Private Sub Command14_Click()
                    On Error Resume Next
   DoCmd.Close
End Sub
Private Sub Command2_Click()
   Dim errorno As Integer, sql As String, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst As New ADODB.Recordset, rst4 As New ADODB.Recordset, RST6 As New ADODB.Recordset
   Dim CON, i As Long, MIAN As Double, MBKM As Double, MOGUDI As Double
   Dim cn As New ADODB.Connection, cm As New ADODB.Command
 ' Dim temp As Double
                     On Error Resume Next
   rst3.Open "INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If IsNull(Me.co) Or Me.co = "" Then
      rst.Open "SELECT TOP 100 PERCENT dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR, dbo.STUF_FSK.MOGODI_A, dbo.STUF_FSK.FI_A, dbo.STUF_FSK.MABL_A FROM  dbo.STUF_DEF INNER JOIN   dbo.STUF_FSK ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_DEF.RADAH =  " & Me.RADAH & " ) ORDER BY dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Else
      rst.Open "SELECT TOP 100 PERCENT dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR, dbo.STUF_FSK.MOGODI_A, dbo.STUF_FSK.FI_A, dbo.STUF_FSK.MABL_A FROM  dbo.STUF_DEF INNER JOIN   dbo.STUF_FSK ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_DEF.RADAH =  " & Me.RADAH & " and dbo.STUF_FSK.CODE = '" & Me.co & "') ORDER BY dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   End If
   RST6.Open "select count(id) from invo_lst where tag <> 20 and tag <> 23", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.Text23 = RST6.Fields(0)
   Set RST6 = New ADODB.Recordset
   RST6.Open "ANBGRD_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   cm.ActiveConnection = CurrentProject.Connection
   cm.CommandTimeout = 3600
   cm.CommandText = "SELECT     TOP 100 PERCENT DATE_N, TAG, NUMBER, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID,  MEGH_R , RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, ID, BARGAH FROM dbo.TMPAV101 ORDER BY DATE_N, BARGAH"
   If MID(Forms![BASEKNOW]![OPTIONSS], 66, 1) = "5" Then
                While Not rst.EOF
                    Me.co = rst.Fields("CODE")
                    MBKM = rst.Fields("MABL_A")
                    MIAN = rst.Fields("FI_A")
                    If MIAN = 0 Then
                        MIAN = GETSTANDARDPRICE(rst.Fields("CODE"))
                        If MIAN = 0 Then
                            MIAN = GETFIRSTPRICE(rst.Fields("CODE"))
                        End If
                    End If
                    MOGUDI = rst.Fields("MOGODI_A")
                    CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS  WHERE TABLE_NAME = 'TMPAV101')   DROP VIEW TMPAV101")
                    CurrentProject.Connection.Execute ("CREATE VIEW  TMPAV101 as " & _
                      " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, dbo.INVO_LST.TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,  dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO,dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM  dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG = dbo.TAGCOD.CODE WHERE  (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH , dbo.INVO_LST.NUMBER UNION " & _
                      " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, 6 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBARF AS ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG + 1  = dbo.TAGCOD.CODE " & _
                      " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBARF = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & Me.dt & ")  AND (dbo.INVO_LST.TAG = 5) ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
                      " SELECT     TOP 100 PERCENT dbo.HEAD_LST_FBK.DATE_N, 4 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN       dbo.HEAD_LST_FBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_FBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_FBK.dtag INNER JOIN  dbo.TAGCOD ON dbo.HEAD_LST_FBK.htag = dbo.TAGCOD.CODE " & _
                      " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_FBK.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST_FBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
                      " SELECT     TOP 100 PERCENT dbo.HEAD_LST_KBK.DATE_N, 3 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A,dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE, dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH  FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST_KBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_KBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_KBK.dtag INNER JOIN " & _
                      " dbo.TAGCOD ON dbo.HEAD_LST_KBK.htag = dbo.TAGCOD.CODE WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_KBK.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST_KBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER  union " & _
                      " SELECT     TOP 100 PERCENT dbo.ANBGRD_HEAD.GRD_DATE, dbo.UIIF(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3, N'>', 0, 18, 17) AS TAG,dbo.ANBGRD_LST.GRD_NUM, dbo.ANBGRD_HEAD.GRD_ANBAR, 1 AS radif, dbo.ANBGRD_LST.CODE,(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS MEG, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS MEGK, 0 AS megh_mar,' ' AS mol, dbo.ANBGRD_LST.MABL, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL AS MABLK, 0 AS froma, '' AS nrasid, 0 AS MEGH_R, dbo.STUF_DEF.RADAH, 0 AS sanadno, '  ' AS cust_no, 0 AS anbarf, dbo.STUF_DEF.VAHED, 0 AS n_kol, 0 AS n_moin, 0 AS n_taf, dbo.ANBGRD_LST.MABL AS avrage, 0 AS id, '17' AS Expr1 FROM   dbo.ANBGRD_LST INNER JOIN  dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM INNER JOIN  dbo.STUF_DEF ON dbo.ANBGRD_LST.CODE = dbo.STUF_DEF.CODE " & _
                      " WHERE      (dbo.ANBGRD_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.ANBGRD_HEAD.GRD_ANBAR = " & rst.Fields("ANBAR") & ") and ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1 <> 0) AND (dbo.ANBGRD_HEAD.GRD_DATE > " & Me.dt & ") AND (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) ORDER BY dbo.ANBGRD_HEAD.GRD_DATE, dbo.ANBGRD_LST.GRD_NUM ")
                    Set RST2 = New ADODB.Recordset
                    Set RST2 = cm.Execute
                    While Not RST2.EOF
                         Text19 = Text19 + 1
                         Me.Repaint
                         rst3.Filter = "ID = " & RST2.Fields("ID")
                         Set rst3 = rst3
                         If RST2.Fields("ID") = 3765 Then
                         DoCmd.Beep
                         End If
                        ' If MIAN < 0 Then
                        '          MIAN = 0
                       '  End If
                         Select Case RST2.Fields("TAG")
                             Case 1: 'خريد
                                  MBKM = MBKM + RST2.Fields("MABL_K")
                                  MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
                                  If MBKM = 0 Then
                               '     MIAN = 0
                                  Else
                                     If MOGUDI = 0 Then
                                '         MIAN = 0
                                         MBKM = 0
                                     Else
                                         MIAN = MBKM / MOGUDI
                                     End If
                                  End If
                                  rst3.Fields("AVRAGE") = MIAN
                                  rst3.update
                             Case 22: 'برگشت فروش سال قبل
                                  If MBKM <= 0 Then
                                      MBKM = RST2.Fields("MABL") * RST2.Fields("MEGH_MAR")
                                  Else
                                    MBKM = MBKM + MIAN * RST2.Fields("MEGH_MAR")
                                  End If
                                  MOGUDI = MOGUDI + RST2.Fields("MEGH_MAR")
                                  If MBKM = 0 Then
                                 '    MIAN = 0
                                  Else
                                     If MOGUDI = 0 Then
                                  '       MIAN = 0
                                         MBKM = 0
                                     Else
                                         MIAN = MBKM / MOGUDI
                                     End If
                                  End If
                                  rst3.Fields("AVRAGE") = MIAN
                                  rst3.update
                             Case 24: 'برگشت فروش
                                  If MBKM <= 0 Then
                                      MBKM = RST2.Fields("MABL_k")
                                  Else
                                      MBKM = MBKM + RST2.Fields("MEGHK") * MIAN
                                  End If
                                  MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
                                  If MBKM = 0 Then
                                   '  MIAN = 0
                                  Else
                                     If MOGUDI = 0 Then
                                    '     MIAN = 0
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
                                     'MIAN = 0
                                  Else
                                     If MOGUDI = 0 Then
                                      '   MIAN = 0
                                         MBKM = 0
                                     Else
                                         MIAN = MBKM / MOGUDI
                                     End If
                                  End If
                                  rst3.Fields("AVRAGE2") = MIAN
                                  rst3.update
                             Case 4: 'برگشت فروش
                                  MBKM = MBKM + (RST2.Fields("MEGH_MAR") * rst3.Fields("AVRAGE"))
                                  MOGUDI = MOGUDI + RST2.Fields("MEGH_MAR")
                                  If MBKM = 0 Then
                                     'MIAN = 0
                                  Else
                                     If MOGUDI = 0 Then
                                      '   MIAN = 0
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
                                  MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
                                  If MBKM = 0 Then
                                     'MIAN = 0
                                  Else
                                     If MOGUDI = 0 Then
                                      '   MIAN = 0
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
                            Case 9:    'توليد
                                  If RST2.Fields("n_kol") <> 0 And Not IsNull(RST2.Fields("n_kol")) And MID(Forms![BASEKNOW]![OPTIONSS], 56, 1) = "5" Then
                                      Set rst4 = New ADODB.Recordset
                                      rst4.Open "SELECT  dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, SUM(dbo.DTL_MANF.MABLK) AS MABLKs FROM         dbo.HEAD_MANF INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE (dbo.HEAD_MANF.FNUMB = " & RST2.Fields("n_kol") & ") GROUP BY dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                                      If rst4.RecordCount > 0 Then
                                         rst3.Fields("MABL") = rst4.Fields(0) + rst4.Fields(1) + rst4.Fields(2)
                                         rst3.Fields("MABL_k") = Round((rst4.Fields(0) + rst4.Fields(1) + rst4.Fields(2)) * RST2.Fields("MEGHK"))
                                         rst3.update
                                      End If
                                   Else
                                        rst3.Fields("MABL") = GETSTANDARDPRICE_KOL(rst.Fields("CODE"), RST2.Fields("date_n"))
                                        If rst3.Fields("MABL") = 0 Then
                                            rst3.Fields("MABL") = GETFIRSTPRICE(rst.Fields("CODE"))
                                        End If
                                        rst3.Fields("MABL_k") = Round(rst3.Fields("MABL") * RST2.Fields("MEGHK"))
                                        rst3.update
                                   End If
                                   MBKM = MBKM + rst3.Fields("MABL_k")
                                   MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
                                   If MBKM = 0 Then
                                     'MIAN = 0
                                   Else
                                     If MOGUDI = 0 Then
                                        ' MIAN = 0
                                         MBKM = 0
                                     Else
                                         MIAN = MBKM / MOGUDI
                                     End If
                                   End If
                                   rst3.Fields("AVRAGE") = MIAN
                                   rst3.update
                             Case 17: 'كسري انبار
                                  MBKM = MBKM + MIAN * RST2.Fields("MEGHK")
                                  MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
                                  If MBKM = 0 Then
                                   '  MIAN = 0
                                  Else
                                     If MOGUDI = 0 Then
                                   '      MIAN = 0
                                         MBKM = 0
                                    Else
                                        MIAN = MBKM / MOGUDI
                                     End If
                                  End If
                                   ' If MIAN < 0 Then
                                    '   MIAN = 0
                                    'End If
                                    RST6.Filter = "code = '" & RST2.Fields("code") & "' and GRD_NUM = " & RST2.Fields("NUMBER")
                                    Set RST6 = RST6
                                    RST6.Fields("MABL") = MIAN
                                    RST6.update
                             Case 18: 'فروش
                                  MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                                  MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                                  RST6.Filter = "code = '" & RST2.Fields("code") & "' and GRD_NUM = " & RST2.Fields("NUMBER")
                                  Set RST6 = RST6
                                  RST6.Fields("MABL") = MIAN
                                  RST6.update
                             Case 26: 'برگشت خريد
                                  MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                                  MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                                  rst3.Fields("AVRAGE") = MIAN
                                  rst3.update
                         End Select
                         RST2.MoveNext
                      Wend
                        RST2.Close
                      rst.MoveNext
                      Me.Text23 = Me.Text23 - 1
                Wend

   Else

       Select Case Me.RADAH
            Case "2":
                    While Not rst.EOF
                        Me.co = rst.Fields("CODE")
                        MIAN = GETSTANDARDPRICE(rst.Fields("CODE"))
                        If MIAN = 0 Then
                            MIAN = GETFIRSTPRICE(rst.Fields("CODE"))
                        End If
                        If MIAN = 0 Then
                              DoCmd.RunSQL "UPDATE    dbo.INVO_LST SET  AVRAGE = 0, AVRAGE2 = 0 WHERE     (CODE = N'" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ")"
                              DoCmd.RunSQL "update dbo.INVO_LST SET MABL = 0, MABL_K = 0 WHERE    ((TAG = 5) OR (TAG = 6) OR (TAG = 7) OR (TAG = 8) OR (TAG = 9) OR (TAG = 10) OR (TAG = 11) OR (TAG = 16) OR (TAG = 17) OR (TAG = 18)) AND ((CODE = N'" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & "))"
                        Else
                            CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS  WHERE TABLE_NAME = 'TMPAV101')   DROP VIEW TMPAV101")
                            CurrentProject.Connection.Execute ("CREATE VIEW  TMPAV101 as " & _
                              " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, dbo.INVO_LST.TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,  dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO,dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM  dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG = dbo.TAGCOD.CODE WHERE  (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH , dbo.INVO_LST.NUMBER UNION " & _
                              " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, 6 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBARF AS ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG + 1 = dbo.TAGCOD.CODE " & _
                              " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBARF = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & Me.dt & ")  AND (dbo.INVO_LST.TAG = 5) ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
                              " SELECT     TOP 100 PERCENT dbo.HEAD_LST_FBK.DATE_N, 4 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN       dbo.HEAD_LST_FBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_FBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_FBK.dtag INNER JOIN  dbo.TAGCOD ON dbo.HEAD_LST_FBK.htag = dbo.TAGCOD.CODE " & _
                              " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_FBK.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST_FBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
                              " SELECT     TOP 100 PERCENT dbo.HEAD_LST_KBK.DATE_N, 3 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A,dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE, dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH  FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST_KBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_KBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_KBK.dtag INNER JOIN " & _
                              " dbo.TAGCOD ON dbo.HEAD_LST_KBK.htag = dbo.TAGCOD.CODE WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_KBK.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST_KBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER union " & _
                      " SELECT     TOP 100 PERCENT dbo.ANBGRD_HEAD.GRD_DATE, dbo.UIIF(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3, N'>', 0, 18, 17) AS TAG,dbo.ANBGRD_LST.GRD_NUM, dbo.ANBGRD_HEAD.GRD_ANBAR, 1 AS radif, dbo.ANBGRD_LST.CODE,(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS MEG, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS MEGK, 0 AS megh_mar,' ' AS mol, dbo.ANBGRD_LST.MABL, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL AS MABLK, 0 AS froma, '' AS nrasid, 0 AS MEGH_R, dbo.STUF_DEF.RADAH, 0 AS sanadno, '  ' AS cust_no, 0 AS anbarf, dbo.STUF_DEF.VAHED, 0 AS n_kol, 0 AS n_moin, 0 AS n_taf, dbo.ANBGRD_LST.MABL AS avrage, 0 AS id, '17' AS Expr1 FROM   dbo.ANBGRD_LST INNER JOIN  dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM INNER JOIN  dbo.STUF_DEF ON dbo.ANBGRD_LST.CODE = dbo.STUF_DEF.CODE " & _
                      " WHERE      (dbo.ANBGRD_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.ANBGRD_HEAD.GRD_ANBAR = " & rst.Fields("ANBAR") & ") and ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1 <> 0) AND (dbo.ANBGRD_HEAD.GRD_DATE > " & Me.dt & ") AND (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) ORDER BY dbo.ANBGRD_HEAD.GRD_DATE, dbo.ANBGRD_LST.GRD_NUM ")
                            Set RST2 = New ADODB.Recordset
                            Set RST2 = cm.Execute
                            While Not RST2.EOF
                                      MIAN = GETSTANDARDPRICE_KOL(rst.Fields("CODE"), RST2.Fields("DATE_N"))
                                 Text19 = Text19 + 1
                                 Me.Repaint
                                 rst3.Filter = "ID = " & RST2.Fields("ID")
                                 Set rst3 = rst3
                                ' If MIAN < 0 Then
                                      '    MIAN = 0
                               ' End If
                                 Select Case RST2.Fields("TAG")
                                     Case 1: 'خريد
                                          rst3.Fields("AVRAGE") = MIAN
                                          rst3.Fields("AVRAGE2") = MIAN
                                          rst3.update
                                     Case 22: 'برگشت فروش سال قبل
                                          rst3.Fields("AVRAGE") = MIAN
                                          rst3.Fields("AVRAGE2") = MIAN
                                          rst3.update
                                     Case 24: 'برگشت فروش سال قبل
                                          rst3.Fields("AVRAGE2") = MIAN
                                          rst3.Fields("AVRAGE") = MIAN
                                          rst3.update
                                     Case 2: 'فروش
                                          rst3.Fields("AVRAGE") = MIAN
                                          rst3.Fields("AVRAGE2") = MIAN
                                          rst3.update
                                     Case 5: 'انتقالي خروج
                                          rst3.Fields("AVRAGE") = MIAN
                                          rst3.Fields("AVRAGE2") = MIAN
                                          rst3.Fields("MABL") = MIAN
                                          rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                          rst3.update
                                     Case 6: 'انتقالي خروج
                                          rst3.Fields("AVRAGE2") = MIAN
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
                                     Case 9:    'tolid
                                          If RST2.Fields("n_kol") <> 0 And Not IsNull(RST2.Fields("n_kol")) And MID(Forms![BASEKNOW]![OPTIONSS], 56, 1) = "5" Then
                                            Set rst4 = New ADODB.Recordset
                                            rst4.Open "SELECT  dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, SUM(dbo.DTL_MANF.MABLK) AS MABLKs FROM         dbo.HEAD_MANF INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE (dbo.HEAD_MANF.FNUMB = " & RST2.Fields("n_kol") & ") GROUP BY dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                                            If rst4.RecordCount > 0 Then
                                               MIAN = rst4.Fields(0) + rst4.Fields(1) + rst4.Fields(2)
                                            End If
                                          End If
                                          rst3.Fields("AVRAGE") = MIAN
                                          rst3.Fields("MABL") = MIAN
                                          rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                          rst3.update
                                    Case 17: 'كسري انبار
                                        ' MBKM = MBKM + rst2.Fields("MABL_K")
                                         RST6.Filter = "code = '" & RST2.Fields("code") & "' and GRD_NUM = " & RST2.Fields("NUMBER")
                                         Set RST6 = RST6
                                         MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
                                         RST6.Fields("MABL") = MIAN
                                         RST6.update
                                    Case 18: 'فروش
                                        ' MBKM = MBKM - (rst2.Fields("MEGHK") * MIAN)
                                         MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                                         RST6.Filter = "code = '" & RST2.Fields("code") & "' and GRD_NUM = " & RST2.Fields("NUMBER")
                                         Set RST6 = RST6
                                         RST6.Fields("MABL") = MIAN
                                         RST6.update
                                     Case 26:    'موادساير خروج
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
                        If MIAN = 0 Then
                            MIAN = GETFIRSTPRICE(rst.Fields("CODE"))
                        End If
                        If MIAN = 0 Then
                              DoCmd.RunSQL "UPDATE    dbo.INVO_LST SET  AVRAGE = 0, AVRAGE2 = 0 WHERE     (CODE = N'" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ")"
                              DoCmd.RunSQL "update dbo.INVO_LST SET MABL = 0, MABL_K = 0 WHERE    ((TAG = 5) OR (TAG = 6) OR (TAG = 7) OR (TAG = 8) OR (TAG = 9) OR (TAG = 10) OR (TAG = 11) OR (TAG = 16) OR (TAG = 17) OR (TAG = 18)) AND ((CODE = N'" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & "))"
                        Else
                            CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS  WHERE TABLE_NAME = 'TMPAV101')   DROP VIEW TMPAV101")
                            CurrentProject.Connection.Execute ("CREATE VIEW  TMPAV101 as " & _
                              " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, dbo.INVO_LST.TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,  dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO,dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM  dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG = dbo.TAGCOD.CODE WHERE  (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH , dbo.INVO_LST.NUMBER UNION " & _
                              " SELECT     TOP 100 PERCENT dbo.HEAD_LST.DATE_N, 6 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBARF AS ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.TAGCOD ON dbo.HEAD_LST.TAG + 1 = dbo.TAGCOD.CODE " & _
                              " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBARF = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST.DATE_N > " & Me.dt & ")  AND (dbo.INVO_LST.TAG = 5) ORDER BY dbo.HEAD_LST.DATE_N,  dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
                              " SELECT     TOP 100 PERCENT dbo.HEAD_LST_FBK.DATE_N, 4 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF,dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL,dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE , dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH FROM         dbo.INVO_LST INNER JOIN       dbo.HEAD_LST_FBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_FBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_FBK.dtag INNER JOIN  dbo.TAGCOD ON dbo.HEAD_LST_FBK.htag = dbo.TAGCOD.CODE " & _
                              " WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_FBK.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST_FBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER UNION " & _
                              " SELECT     TOP 100 PERCENT dbo.HEAD_LST_KBK.DATE_N, 3 AS TAG, dbo.INVO_LST.NUMBER, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A,dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE, dbo.INVO_LST.ID, dbo.TAGCOD.BARGAH  FROM         dbo.INVO_LST INNER JOIN   dbo.HEAD_LST_KBK ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST_KBK.NUMBER1 AND dbo.INVO_LST.TAG = dbo.HEAD_LST_KBK.dtag INNER JOIN " & _
                              " dbo.TAGCOD ON dbo.HEAD_LST_KBK.htag = dbo.TAGCOD.CODE WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_KBK.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST_KBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER union " & _
                      " SELECT     TOP 100 PERCENT dbo.ANBGRD_HEAD.GRD_DATE, dbo.UIIF(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3, N'>', 0, 18, 17) AS TAG,dbo.ANBGRD_LST.GRD_NUM, dbo.ANBGRD_HEAD.GRD_ANBAR, 1 AS radif, dbo.ANBGRD_LST.CODE,(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS MEG, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS MEGK, 0 AS megh_mar,' ' AS mol, dbo.ANBGRD_LST.MABL, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL AS MABLK, 0 AS froma, '' AS nrasid, 0 AS MEGH_R, dbo.STUF_DEF.RADAH, 0 AS sanadno, '  ' AS cust_no, 0 AS anbarf, dbo.STUF_DEF.VAHED, 0 AS n_kol, 0 AS n_moin, 0 AS n_taf, dbo.ANBGRD_LST.MABL AS avrage, 0 AS id, '17' AS Expr1 FROM   dbo.ANBGRD_LST INNER JOIN  dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM INNER JOIN  dbo.STUF_DEF ON dbo.ANBGRD_LST.CODE = dbo.STUF_DEF.CODE " & _
                      " WHERE      (dbo.ANBGRD_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.ANBGRD_HEAD.GRD_ANBAR = " & rst.Fields("ANBAR") & ") and ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1 <> 0) AND (dbo.ANBGRD_HEAD.GRD_DATE > " & Me.dt & ") AND (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) ORDER BY dbo.ANBGRD_HEAD.GRD_DATE, dbo.ANBGRD_LST.GRD_NUM ")
                        Set RST2 = New ADODB.Recordset
                        Set RST2 = cm.Execute
                        While Not RST2.EOF
                                      MIAN = GETSTANDARDPRICE_KOL(rst.Fields("CODE"), RST2.Fields("DATE_N"))
                                 Text19 = Text19 + 1
                                 Me.Repaint
                                 rst3.Filter = "ID = " & RST2.Fields("ID")
                                 Set rst3 = rst3
                             '    If MIAN < 0 Then
                              '            MIAN = 0
                              '  End If
                                 Select Case RST2.Fields("TAG")
                                     Case 1: 'خريد
                                          rst3.Fields("AVRAGE") = MIAN
                                          rst3.Fields("AVRAGE2") = MIAN
                                          rst3.update
                                     Case 22: 'برگشت فروش سال قبل
                                          rst3.Fields("AVRAGE") = MIAN
                                          rst3.Fields("AVRAGE2") = MIAN
                                          rst3.update
                                     Case 24: 'برگشت فروش سال قبل
                                          rst3.Fields("AVRAGE") = MIAN
                                          rst3.Fields("AVRAGE2") = MIAN
                                          rst3.update
                                     Case 2: 'فروش
                                          rst3.Fields("AVRAGE") = MIAN
                                          rst3.Fields("AVRAGE2") = MIAN
                                          rst3.update
                                     Case 6: 'انتقالي خروج
                                          rst3.Fields("AVRAGE") = MIAN
                                          rst3.Fields("AVRAGE2") = MIAN
                                          rst3.Fields("MABL") = MIAN
                                          rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                          rst3.update
                                     Case 5: 'انتقالي خروج
                                          rst3.Fields("AVRAGE") = MIAN
                                          rst3.Fields("AVRAGE2") = MIAN
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
                                          If RST2.Fields("n_kol") <> 0 And Not IsNull(RST2.Fields("n_kol")) And MID(Forms![BASEKNOW]![OPTIONSS], 56, 1) = "5" Then
                                            Set rst4 = New ADODB.Recordset
                                            rst4.Open "SELECT     dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, SUM(dbo.DTL_MANF.MABLK) AS MABLKs FROM         dbo.HEAD_MANF INNER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE (dbo.HEAD_MANF.FNUMB = " & RST2.Fields("n_kol") & ") GROUP BY dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                                            If rst4.RecordCount > 0 Then
                                               MIAN = rst4.Fields(0) + rst4.Fields(1) + rst4.Fields(2)
                                            End If
                                          End If
                                          rst3.Fields("AVRAGE") = MIAN
                                          rst3.Fields("MABL") = MIAN
                                          rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                          rst3.update
                                    Case 17: 'كسري انبار
                                        ' MBKM = MBKM + rst2.Fields("MABL_K")
                                         RST6.Filter = "code = '" & RST2.Fields("code") & "' and GRD_NUM = " & RST2.Fields("NUMBER")
                                         Set RST6 = RST6
                                         MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
                                         RST6.Fields("MABL") = MIAN
                                         RST6.update
                                    Case 18: 'فروش
                                        ' MBKM = MBKM - (rst2.Fields("MEGHK") * MIAN)
                                         MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                                         RST6.Filter = "code = '" & RST2.Fields("code") & "' and GRD_NUM = " & RST2.Fields("NUMBER")
                                         Set RST6 = RST6
                                         RST6.Fields("MABL") = MIAN
                                         RST6.update
                                     Case 26:    'موادساير خروج
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
                      " dbo.TAGCOD ON dbo.HEAD_LST_KBK.htag = dbo.TAGCOD.CODE WHERE     (dbo.INVO_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.INVO_LST.ANBAR = " & rst.Fields("ANBAR") & ") AND (dbo.HEAD_LST_KBK.DATE_N > " & Me.dt & ") ORDER BY dbo.HEAD_LST_KBK.DATE_N, dbo.TAGCOD.BARGAH, dbo.INVO_LST.NUMBER  union " & _
                      " SELECT     TOP 100 PERCENT dbo.ANBGRD_HEAD.GRD_DATE, dbo.UIIF(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3, N'>', 0, 18, 17) AS TAG,dbo.ANBGRD_LST.GRD_NUM, dbo.ANBGRD_HEAD.GRD_ANBAR, 1 AS radif, dbo.ANBGRD_LST.CODE,(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS MEG, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS MEGK, 0 AS megh_mar,' ' AS mol, dbo.ANBGRD_LST.MABL, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL AS MABLK, 0 AS froma, '' AS nrasid, 0 AS MEGH_R, dbo.STUF_DEF.RADAH, 0 AS sanadno, '  ' AS cust_no, 0 AS anbarf, dbo.STUF_DEF.VAHED, 0 AS n_kol, 0 AS n_moin, 0 AS n_taf, dbo.ANBGRD_LST.MABL AS avrage, 0 AS id, '17' AS Expr1 FROM   dbo.ANBGRD_LST INNER JOIN  dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM INNER JOIN  dbo.STUF_DEF ON dbo.ANBGRD_LST.CODE = dbo.STUF_DEF.CODE " & _
                      " WHERE      (dbo.ANBGRD_LST.CODE = '" & rst.Fields("CODE") & "') AND (dbo.ANBGRD_HEAD.GRD_ANBAR = " & rst.Fields("ANBAR") & ") and ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1 <> 0) AND (dbo.ANBGRD_HEAD.GRD_DATE > " & Me.dt & ") AND (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) ORDER BY dbo.ANBGRD_HEAD.GRD_DATE, dbo.ANBGRD_LST.GRD_NUM ")
                    Set RST2 = New ADODB.Recordset
                  Set RST2 = cm.Execute
                    While Not RST2.EOF
                         Text19 = Text19 + 1
                         Me.Repaint
                         rst3.Filter = "ID = " & RST2.Fields("ID")
                         Set rst3 = rst3
                         If RST2.Fields("ID") = 9688 Then
                         DoCmd.Beep
                         End If
                        ' If MIAN < 0 Then
                        '          MIAN = 0
                       '  End If
                         Select Case RST2.Fields("TAG")
                             Case 1: 'خريد
                                  MBKM = MBKM + RST2.Fields("MABL_K")
                                  MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
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
                                  MBKM = MBKM + MIAN * RST2.Fields("MEGH_MAR")
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
                             Case 24: 'برگشت فروش سال قبل
                                  MBKM = MBKM + RST2.Fields("MEGHK") * MIAN
                                  MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
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
                                  MBKM = MBKM + (RST2.Fields("MEGH_MAR") * rst3.Fields("AVRAGE"))
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
                                  MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
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
                            Case 9:    'توليد
                                  MBKM = MBKM + RST2.Fields("MABL_K")
                                  MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
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
                                 rst3.Fields("MABL") = MIAN
                                 rst3.Fields("MABL_k") = Round(MIAN * RST2.Fields("MEGHK"))
                                 rst3.update
                             Case 17: 'كسري انبار
                                  MBKM = MBKM + MIAN * RST2.Fields("MEGHK")
                                  MOGUDI = Round(MOGUDI + RST2.Fields("MEGHK"), 6)
                                  If MBKM = 0 Then
                                     MIAN = 0
                                  Else
                                     If MOGUDI = 0 Then
                                         MIAN = 0
                                         MBKM = 0
      '                               Else
       '                                  MIAN = MBKM / MOGUDI
                                     End If
                                  End If
                                If MIAN < 0 Then
                                   MIAN = 0
                                End If
                                RST6.Filter = "code = '" & RST2.Fields("code") & "' and GRD_NUM = " & RST2.Fields("NUMBER")
                                Set RST6 = RST6
                                RST6.Fields("MABL") = MIAN
                                RST6.update
                             Case 18: 'فروش
                                  MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                                  MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                                RST6.Filter = "code = '" & RST2.Fields("code") & "' and GRD_NUM = " & RST2.Fields("NUMBER")
                                Set RST6 = RST6
                                RST6.Fields("MABL") = MIAN
                                RST6.update
                             Case 26: 'برگشت خريد
                                  MBKM = MBKM - (RST2.Fields("MEGHK") * MIAN)
                                  MOGUDI = MOGUDI - RST2.Fields("MEGHK")
                                  rst3.Fields("AVRAGE") = MIAN
                                  rst3.update
                         End Select
                         RST2.MoveNext
                      Wend
                        RST2.Close
                      rst.MoveNext
                Wend

       End Select
    End If
    rst.Close
    If Err = 0 Then
        DoCmd.Hourglass False
        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات بازسازي  موجودي انبارها با موفقيت انجام شد"
        Me.Text0.Visible = False
        Me.Box7.Visible = False
        Me.Refresh
        DoCmd.Close acForm, Me.NAME
    End If

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
