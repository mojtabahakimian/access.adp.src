CREATE FUNCTION [dbo].[Q_GARDESH_KHFR_DAFTAR_SUB]
   (@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 bigint,
   @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2 bigint,
   @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF nvarchar(50)
   )
   RETURNS TABLE
   AS
   RETURN ( SELECT     dbo.UDATEADD(dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.MAS) AS SDATE, dbo.HEAD_LST.NUMBER AS N_S, dbo.HEAD_LST.CUST_NO, dbo.CUST_HESAB.NAME, 
                         dbo.STUF_DEF.NAME + ' - ' +ISNULL(dbo.INVO_LST.MANDAH,' ' ) + ' - ' + ISNULL(dbo.HEAD_LST.MOLAH,' ' ) AS SHARH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.DATE_N, 
                         dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR AS MEGK, dbo.INVO_LST.MABL, 0 AS bes, (dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) 
                         * dbo.INVO_LST.MABL AS bed, dbo.INVO_LST.RADIF, dbo.INVO_LST.NUMBER
   FROM         dbo.CUST_HESAB INNER JOIN
                         dbo.HEAD_LST INNER JOIN
                         dbo.STUF_DEF INNER JOIN
                         dbo.INVO_LST ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND 
                         dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER ON dbo.CUST_HESAB.hes = dbo.HEAD_LST.CUST_NO
   WHERE     (dbo.HEAD_LST.DATE_N BETWEEN @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 AND @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2) AND 
                         (dbo.HEAD_LST.TAG = 2) AND (dbo.HEAD_LST.CUST_NO = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF)
   UNION
   SELECT     dbo.UDATEADD(dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.MAS) AS SDATE, dbo.HEAD_LST.NUMBER AS N_S, dbo.HEAD_LST.CUST_NO, dbo.CUST_HESAB.NAME, 
                         dbo.STUF_DEF.NAME + ' - ' + ISNULL(dbo.INVO_LST.MANDAH,' ' ) + ' - ' + ISNULL(dbo.HEAD_LST.MOLAH,' ') AS SHARH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.DATE_N, 
                         dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR AS MEGK, dbo.INVO_LST.MABL, (dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) 
                         * dbo.INVO_LST.MABL AS bes, 0 AS bed, dbo.INVO_LST.RADIF, dbo.INVO_LST.NUMBER
   FROM         dbo.CUST_HESAB INNER JOIN
                         dbo.HEAD_LST INNER JOIN
                         dbo.STUF_DEF INNER JOIN
                         dbo.INVO_LST ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND 
                         dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER ON dbo.CUST_HESAB.hes = dbo.HEAD_LST.CUST_NO
   WHERE     (dbo.HEAD_LST.DATE_N BETWEEN @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 AND @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2) AND 
                         (dbo.HEAD_LST.TAG = 1) AND (dbo.HEAD_LST.CUST_NO = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF)
   UNION
   SELECT     ISNULL(dbo.PAY_GETD.DATE_S, dbo.DEED_HED.DATE_S) AS SDATE, dbo.DEED_HED.N_S, RTRIM(CAST(dbo.DEED_DTL.HES_K AS nvarchar)) 
                      + '-' + RTRIM(CAST(dbo.DEED_DTL.HES_M AS nvarchar)) + '-' + RTRIM(CAST(dbo.DEED_DTL.HES_T AS nvarchar)) AS HES, dbo.TDETA_HES.NAME, 
                      dbo.DEED_DTL.SHARH, 0 AS mas, dbo.DEED_HED.DATE_S AS SARDATE, 0 AS MEGHk, dbo.DEED_DTL.BES + dbo.DEED_DTL.BED AS mabl, 
                      dbo.DEED_DTL.BES, dbo.DEED_DTL.BED, dbo.DEED_DTL.RADIF, dbo.TDETA_HES.TNUMBER
FROM         dbo.TDETA_HES INNER JOIN
                      dbo.DEED_HED INNER JOIN
                      dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S ON dbo.TDETA_HES.TNUMBER = dbo.DEED_DTL.HES_T AND 
                      dbo.TDETA_HES.NUMBER = dbo.DEED_DTL.HES_M AND dbo.TDETA_HES.N_KOL = dbo.DEED_DTL.HES_K LEFT OUTER JOIN
                      dbo.PAY_GETD ON dbo.DEED_DTL.N_SERI = dbo.PAY_GETD.N_SERI AND dbo.DEED_DTL.BANK = dbo.PAY_GETD.BANK
WHERE     (RTRIM(CAST(dbo.DEED_DTL.HES_K AS nvarchar)) + '-' + RTRIM(CAST(dbo.DEED_DTL.HES_M AS nvarchar)) 
                      + '-' + RTRIM(CAST(dbo.DEED_DTL.HES_T AS nvarchar)) = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF) AND (dbo.DEED_DTL.RADIF IS NULL) 
                      AND (dbo.DEED_HED.DATE_S BETWEEN @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 AND @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2)
   UNION
   SELECT     DATE_S, 0 AS mas, HES, NAME, SHARH, MAS AS Expr3, DATE_S AS Expr1, 0 AS MEGHk, MABL, dbo.UIIF(MAND, '>=', 0, 0, ABS(MAND)) AS Expr4, 
                         dbo.UIIF(MAND, '>=', 0, MAND, 0) AS Expr5, Expr1 AS Expr2, 0 AS number
   FROM         dbo.Q_GARDESH_KHFR_MAND(@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1, @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF) 
                        Q_GARDESH_KHFR_MAND)
