CREATE FUNCTION dbo.VISITOR_DTL_KALA
									(
									    @dt1 bigint,
									    @dt2 bigint,
									    @visitor nvarchar(40)
									)
									RETURNS TABLE
									AS
									RETURN
									(
									    SELECT TOP (100) PERCENT
									           il.CODE,
									           SUM(il.MEGHk / ISNULL((SELECT COUNT(*) FROM dbo.VISITOR_DTL WHERE NUMBER = il.NUMBER AND TAG = il.TAG), 1))                 AS MEGHk,
									           SUM(il.MABL_K / ISNULL((SELECT COUNT(*) FROM dbo.VISITOR_DTL WHERE NUMBER = il.NUMBER AND TAG = il.TAG), 1))               AS MABL_K,
									           SUM(il.IMBAA)                AS IMBAA,
									           SUM(il.N_MOIN)               AS N_MOIN,
									           sd.NAME                      AS kala,
									           ch.NAME                      AS VISITOR,
									           vd.CUST_NO,
									           SUM(il.MEGH_MAR)             AS MEGH_MAR,
									           SUM(il.MEGH_MAR * il.MABL)   AS MABMAR,
									           SUM(il.MABL_K - il.MEGH_MAR * il.MABL + il.IMBAA - il.N_MOIN) AS GHABEL,
									           ch.ADDRESS,
									           ch.TEL,
									           ch.TOZIH,
									           ch.MOBILE,
									           sd.MENUIT,
									           hl.DEPATMAN                  -- ⭐️ ستون جدید
									    FROM   dbo.HEAD_LST        AS hl
									           INNER JOIN dbo.INVO_LST   AS il ON hl.NUMBER = il.NUMBER AND hl.TAG = il.TAG
									           INNER JOIN dbo.VISITOR_DTL AS vd ON hl.NUMBER = vd.NUMBER AND hl.TAG = vd.TAG
									           INNER JOIN dbo.STUF_DEF    AS sd ON il.CODE   = sd.CODE
									           INNER JOIN dbo.TCOD_VAHEDS AS tv ON il.VAHED_K = tv.CODE
									           INNER JOIN dbo.CUST_HESAB  AS ch ON vd.CUST_NO = ch.hes
									    WHERE  hl.DATE_N BETWEEN @dt1 AND @dt2
									      AND  hl.TAG = 2
									    GROUP BY
									           il.CODE, sd.NAME, ch.NAME, vd.CUST_NO,
									           ch.ADDRESS, ch.TEL, ch.TOZIH, ch.MOBILE,
									           sd.MENUIT, hl.DEPATMAN       -- ⭐️ در GROUP BY هم اضافه شود
									    HAVING vd.CUST_NO LIKE @visitor
									)
