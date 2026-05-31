 
 
  CREATE FUNCTION [dbo].[VISITOR_DTL_KALA_mar]
 
   (@dt1 bigint,
 
   @dt2 bigint,
 
   @visitor nvarchar(40))
 
   RETURNS TABLE
 
   AS
 
   RETURN ( SELECT     TOP 100 PERCENT dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGH_MAR) AS MEGHk, SUM(dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.MABL) 
 
                         AS MABL_K, SUM(dbo.INVO_LST.IMBAA /dbo.UIIF(dbo.INVO_LST.MEGHk * dbo.INVO_LST.MEGH_MAR,'=', 0, 999999999999,  dbo.INVO_LST.MEGHk * dbo.INVO_LST.MEGH_MAR)) AS IMBAA, 
 
                         SUM(dbo.INVO_LST.N_KOL * dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR / 100 + ROUND(ROUND((dbo.INVO_LST.MABL_K - ROUND(dbo.INVO_LST.MABL_K
 
                          * dbo.INVO_LST.N_KOL / 100, 0)) * dbo.INVO_LST.TKHN / 100, 0) / dbo.UIIF(dbo.INVO_LST.MEGHk * dbo.INVO_LST.MEGH_MAR,'=', 0, 999999999999,  dbo.INVO_LST.MEGHk * dbo.INVO_LST.MEGH_MAR), 0)) AS N_MOIN, 
 
                         dbo.STUF_DEF.NAME AS kala, dbo.CUST_HESAB.NAME AS VISITOR, dbo.VISITOR_DTL.CUST_NO, SUM(dbo.INVO_LST.MEGH_MAR) AS MEGH_MAR, 
 
                         SUM(dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.MABL) AS MABMAR, 
 
                         SUM((dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.MABL + dbo.INVO_LST.IMBAA / dbo.UIIF(dbo.INVO_LST.MEGHk * dbo.INVO_LST.MEGH_MAR,'=', 0, 999999999999,  dbo.INVO_LST.MEGHk * dbo.INVO_LST.MEGH_MAR)) 
 
                         - (dbo.INVO_LST.N_KOL * dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR / 100 + ROUND(ROUND((dbo.INVO_LST.MABL_K - ROUND(dbo.INVO_LST.MABL_K
 
                          * dbo.INVO_LST.N_KOL / 100, 0)) * dbo.INVO_LST.TKHN / 100, 0) / dbo.UIIF(dbo.INVO_LST.MEGHk * dbo.INVO_LST.MEGH_MAR,'=', 0, 999999999999,  dbo.INVO_LST.MEGHk * dbo.INVO_LST.MEGH_MAR), 0))) AS GHABEL, 
 
                         dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.TOZIH, dbo.CUST_HESAB.MOBILE, dbo.STUF_DEF.MENUIT
 
   FROM         dbo.HEAD_LST INNER JOIN
 
                         dbo.INVO_LST ON dbo.HEAD_LST.NUMBER1 = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG + 2 INNER JOIN
 
                         dbo.VISITOR_DTL ON dbo.HEAD_LST.NUMBER = dbo.VISITOR_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.VISITOR_DTL.TAG INNER JOIN
 
                         dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN
 
                         dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE INNER JOIN
 
                         dbo.CUST_HESAB ON dbo.VISITOR_DTL.CUST_NO = dbo.CUST_HESAB.hes
 
   WHERE     (dbo.HEAD_LST.DATE_N BETWEEN @dt1 AND @dt2) AND (dbo.HEAD_LST.TAG = 4)
 
   GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.CUST_HESAB.NAME, dbo.VISITOR_DTL.CUST_NO, dbo.CUST_HESAB.ADDRESS, 
 
                         dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.TOZIH, dbo.CUST_HESAB.MOBILE, dbo.STUF_DEF.MENUIT
 
   HAVING      (dbo.VISITOR_DTL.CUST_NO LIKE @visitor) AND (SUM(dbo.INVO_LST.MEGH_MAR) > 0)
 
   ORDER BY dbo.CUST_HESAB.NAME, SUM(dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.MABL) DESC
 
  UNION
 
  SELECT     TOP 100 PERCENT dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk) AS MEGHk, SUM(dbo.INVO_LST.MABL_K) AS MABL_K, 
 
                        SUM(dbo.INVO_LST.IMBAA) AS IMBAA, SUM(dbo.INVO_LST.N_MOIN) AS N_MOIN, dbo.STUF_DEF.NAME AS kala, dbo.CUST_HESAB.NAME AS VISITOR, 
 
                        dbo.VISITOR_DTL.CUST_NO, SUM(dbo.INVO_LST.MEGHk) AS MEGH_MAR, SUM(dbo.INVO_LST.MABL_K) AS MABMAR, 
 
                        SUM(dbo.INVO_LST.MABL_K + dbo.INVO_LST.IMBAA - dbo.INVO_LST.N_MOIN) AS GHABEL, dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, 
 
                        dbo.CUST_HESAB.TOZIH, dbo.CUST_HESAB.MOBILE, dbo.STUF_DEF.MENUIT
 
  FROM         dbo.HEAD_LST INNER JOIN
 
                        dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER INNER JOIN
 
                        dbo.VISITOR_DTL ON dbo.HEAD_LST.NUMBER = dbo.VISITOR_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.VISITOR_DTL.TAG INNER JOIN
 
                        dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN
 
                        dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE INNER JOIN
 
                        dbo.CUST_HESAB ON dbo.VISITOR_DTL.CUST_NO = dbo.CUST_HESAB.hes
 
  WHERE     (dbo.HEAD_LST.DATE_N BETWEEN @dt1 AND @dt2) AND (dbo.HEAD_LST.TAG = 24)
 
  GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.CUST_HESAB.NAME, dbo.VISITOR_DTL.CUST_NO, dbo.CUST_HESAB.ADDRESS, 
 
                        dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.TOZIH, dbo.CUST_HESAB.MOBILE, dbo.STUF_DEF.MENUIT
 
  HAVING      (dbo.VISITOR_DTL.CUST_NO LIKE @visitor) AND (SUM(dbo.INVO_LST.MEGHk) > 0)
 
  ORDER BY dbo.CUST_HESAB.NAME, SUM(dbo.INVO_LST.MABL_K) DESC )
 

