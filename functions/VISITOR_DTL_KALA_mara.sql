  
 
  CREATE FUNCTION [dbo].[VISITOR_DTL_KALA_mara]
 
    (@dt1 bigint,
 
    @dt2 bigint,
 
    @visitor nvarchar(40))
 
    RETURNS TABLE
 
    AS
 
    RETURN ( SELECT     TOP 100 PERCENT dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk / ISNULL((SELECT COUNT(*) FROM dbo.VISITOR_DTL WHERE NUMBER = dbo.INVO_LST.NUMBER AND TAG = dbo.INVO_LST.TAG), 1)) AS MEGHk, SUM(dbo.INVO_LST.MABL_K / ISNULL((SELECT COUNT(*) FROM dbo.VISITOR_DTL WHERE NUMBER = dbo.INVO_LST.NUMBER AND TAG = dbo.INVO_LST.TAG), 1)) AS MABL_K,
 
                        SUM(dbo.INVO_LST.IMBAA) AS IMBAA, SUM(dbo.INVO_LST.N_MOIN) AS N_MOIN, dbo.STUF_DEF.NAME AS kala, dbo.CUST_HESAB.NAME AS VISITOR, 
 
                        dbo.VISITOR_DTL.CUST_NO, ISNULL(VISITOR_DTL_KALA_mar.MEGH_MAR, 0) AS MEGH_MAR, ISNULL(VISITOR_DTL_KALA_mar.GHABEL, 0) 
 
                        AS MABMAR, SUM(dbo.INVO_LST.MABL_K - ISNULL(VISITOR_DTL_KALA_mar.GHABEL, 0) + dbo.INVO_LST.IMBAA - dbo.INVO_LST.N_MOIN) 
 
                        AS GHABEL, dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.TOZIH, dbo.CUST_HESAB.MOBILE, 
 
                        dbo.STUF_DEF.MENUIT
 
  FROM         dbo.HEAD_LST INNER JOIN
 
                        dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN
 
                        dbo.VISITOR_DTL ON dbo.HEAD_LST.NUMBER = dbo.VISITOR_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.VISITOR_DTL.TAG INNER JOIN
 
                        dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN
 
                        dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE INNER JOIN
 
                        dbo.CUST_HESAB ON dbo.VISITOR_DTL.CUST_NO = dbo.CUST_HESAB.hes LEFT OUTER JOIN
 
                        dbo.VISITOR_DTL_KALA_mar(@dt1, @dt2, @visitor) VISITOR_DTL_KALA_mar ON 
 
                        dbo.VISITOR_DTL.CUST_NO = VISITOR_DTL_KALA_mar.CUST_NO AND dbo.INVO_LST.CODE = VISITOR_DTL_KALA_mar.CODE
 
  WHERE     (dbo.HEAD_LST.DATE_N BETWEEN @dt1 AND @dt2) AND (dbo.HEAD_LST.TAG = 2)
 
  GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.CUST_HESAB.NAME, dbo.VISITOR_DTL.CUST_NO, dbo.CUST_HESAB.ADDRESS, 
 
                        dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.TOZIH, dbo.CUST_HESAB.MOBILE, dbo.STUF_DEF.MENUIT, ISNULL(VISITOR_DTL_KALA_mar.MEGH_MAR, 
 
                        0), ISNULL(VISITOR_DTL_KALA_mar.GHABEL, 0)
 
  HAVING      (dbo.VISITOR_DTL.CUST_NO LIKE @visitor)
 
  ORDER BY dbo.CUST_HESAB.NAME, SUM(dbo.INVO_LST.MABL_K) DESC
 
    UNION
 
    SELECT     TOP 100 PERCENT dbo.INVO_LST.CODE, SUM(0) AS MEGHk, SUM(0) AS MABL_K, SUM(0) AS IMBAA, SUM(0) AS N_MOIN, 
 
                          dbo.STUF_DEF.NAME AS kala, dbo.CUST_HESAB.NAME AS VISITOR, dbo.VISITOR_DTL.CUST_NO, SUM(dbo.INVO_LST.MEGH_MAR / ISNULL((SELECT COUNT(*) FROM dbo.VISITOR_DTL WHERE NUMBER = dbo.INVO_LST.NUMBER AND TAG = dbo.INVO_LST.TAG), 1)) AS MEGH_MAR,
 
                          SUM(dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.MABL) AS MABMAR, 
 
                          SUM((dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.MABL + dbo.INVO_LST.IMBAA /dbo.UIIF(dbo.INVO_LST.MEGHk * dbo.INVO_LST.MEGH_MAR,'=', 0, 999999999999,  dbo.INVO_LST.MEGHk * dbo.INVO_LST.MEGH_MAR)) 
 
                          - (dbo.INVO_LST.N_KOL * dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR / 100 + ROUND(ROUND((dbo.INVO_LST.MABL_K - ROUND(dbo.INVO_LST.MABL_K
 
                           * dbo.INVO_LST.N_KOL / 100, 0)) * dbo.INVO_LST.TKHN / 100, 0) / dbo.UIIF(dbo.INVO_LST.MEGHk * dbo.INVO_LST.MEGH_MAR,'=', 0, 999999999999,  dbo.INVO_LST.MEGHk * dbo.INVO_LST.MEGH_MAR), 0))) AS GHABEL, 
 
                          dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.TOZIH, dbo.CUST_HESAB.MOBILE, dbo.STUF_DEF.MENUIT
 
    FROM         dbo.HEAD_LST INNER JOIN
 
                          dbo.INVO_LST ON dbo.HEAD_LST.NUMBER1 = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG + 2 INNER JOIN
 
                          dbo.VISITOR_DTL ON dbo.HEAD_LST.NUMBER = dbo.VISITOR_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.VISITOR_DTL.TAG INNER JOIN
 
                          dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN
 
                          dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE INNER JOIN
 
                          dbo.CUST_HESAB ON dbo.VISITOR_DTL.CUST_NO = dbo.CUST_HESAB.hes LEFT OUTER JOIN
 
                          dbo.VISITOR_DTL_KALA(@dt1, @dt2, @visitor) VISITOR_DTL_KALA ON dbo.INVO_LST.CODE = VISITOR_DTL_KALA.CODE AND 
 
                          dbo.VISITOR_DTL.CUST_NO = VISITOR_DTL_KALA.CUST_NO
 
    WHERE     (dbo.HEAD_LST.DATE_N BETWEEN @dt1 AND @dt2) AND (dbo.HEAD_LST.TAG = 4)
 
    GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.CUST_HESAB.NAME, dbo.VISITOR_DTL.CUST_NO, dbo.CUST_HESAB.ADDRESS, 
 
                          dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.TOZIH, dbo.CUST_HESAB.MOBILE, dbo.STUF_DEF.MENUIT, VISITOR_DTL_KALA.CODE
 
    HAVING      (dbo.VISITOR_DTL.CUST_NO LIKE @visitor) AND (VISITOR_DTL_KALA.CODE IS NULL) AND (SUM(dbo.INVO_LST.MEGH_MAR) > 0)
 
    ORDER BY dbo.CUST_HESAB.NAME, SUM(0) DESC
 
   UNION
 
   SELECT     TOP 100 PERCENT dbo.INVO_LST.CODE, SUM(0) AS MEGHk, SUM(0) AS MABL_K, SUM(0) AS IMBAA, SUM(0) AS N_MOIN, 
 
                         dbo.STUF_DEF.NAME AS kala, dbo.CUST_HESAB.NAME AS VISITOR, dbo.VISITOR_DTL.CUST_NO, SUM(dbo.INVO_LST.MEGHk / ISNULL((SELECT COUNT(*) FROM dbo.VISITOR_DTL WHERE NUMBER = dbo.INVO_LST.NUMBER AND TAG = dbo.INVO_LST.TAG), 1)) AS MEGH_MAR,
 
                         SUM(dbo.INVO_LST.MABL_K / ISNULL((SELECT COUNT(*) FROM dbo.VISITOR_DTL WHERE NUMBER = dbo.INVO_LST.NUMBER AND TAG = dbo.INVO_LST.TAG), 1)) AS MABMAR, SUM(dbo.INVO_LST.MABL_K + dbo.INVO_LST.IMBAA - dbo.INVO_LST.N_MOIN) AS GHABEL,
 
                         dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.TOZIH, dbo.CUST_HESAB.MOBILE, dbo.STUF_DEF.MENUIT
 
   FROM         dbo.CUST_HESAB INNER JOIN
 
                         dbo.HEAD_LST INNER JOIN
 
                         dbo.VISITOR_DTL ON dbo.HEAD_LST.NUMBER = dbo.VISITOR_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.VISITOR_DTL.TAG ON 
 
                         dbo.CUST_HESAB.hes = dbo.VISITOR_DTL.CUST_NO INNER JOIN
 
                         dbo.STUF_DEF INNER JOIN
 
                         dbo.INVO_LST ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE INNER JOIN
 
                         dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND 
 
                         dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG LEFT OUTER JOIN
 
                         dbo.VISITOR_DTL_KALA(@dt1, @dt2, @visitor) VISITOR_DTL_KALA ON dbo.INVO_LST.CODE = VISITOR_DTL_KALA.CODE AND 
 
                         dbo.VISITOR_DTL.CUST_NO = VISITOR_DTL_KALA.CUST_NO
 
   WHERE     (dbo.HEAD_LST.DATE_N BETWEEN @dt1 AND @dt2) AND (dbo.HEAD_LST.TAG = 24)
 
   GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.CUST_HESAB.NAME, dbo.VISITOR_DTL.CUST_NO, dbo.CUST_HESAB.ADDRESS, 
 
                         dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.TOZIH, dbo.CUST_HESAB.MOBILE, dbo.STUF_DEF.MENUIT, VISITOR_DTL_KALA.CODE
 
   HAVING      (dbo.VISITOR_DTL.CUST_NO LIKE @visitor) AND (VISITOR_DTL_KALA.CODE IS NULL)
 
   ORDER BY dbo.CUST_HESAB.NAME, SUM(0) DESC )
 
  

