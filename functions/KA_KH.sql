 
 
 
  
   
    CREATE FUNCTION dbo.KA_KH
        (@Forms___F_MENU_KART___FT BIGINT)
        RETURNS TABLE
        AS
        RETURN ( SELECT     ANBAR, CODE, SUM(MOGODI_A) AS MEG, FI_A AS mabl_a, FI_A * MOGODI_A AS MABLK, 0 AS TAG, MOGODI_A AS MEGK, 
                             @Forms___F_MENU_KART___FT AS DATE_N, NULL AS NUMBER, ' ' AS BEDNAME, FI_A AS avrage, 0 AS FNUMCO, IDD, ' ' AS mol
       FROM         dbo.STUF_FSK
       GROUP BY ANBAR, CODE, FI_A, MOGODI_A, IDD
        UNION
       SELECT     dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGHk AS MEG, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.HEAD_LST.TAG, 
                             dbo.INVO_LST.MEGHk, dbo.HEAD_LST.DATE_N, dbo.INVO_LST.NUMBER, dbo.CUST_HESAB.NAME, dbo.INVO_LST.AVRAGE, dbo.HEAD_LST.FNUMCO, 
                             dbo.INVO_LST.id, ISNULL(dbo.INVO_LST.MANDAH, N'  ') AS mol
       
       FROM         dbo.HEAD_LST INNER JOIN
                             dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER LEFT OUTER JOIN
                             dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes
       WHERE     (dbo.HEAD_LST.TAG = 1) OR
                             (dbo.HEAD_LST.TAG = 7) OR
                             (dbo.HEAD_LST.TAG = 9) OR
                              (dbo.HEAD_LST.TAG = 24)
   
        UNION
  SELECT     dbo.INVO_LST.ANBARF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGHk AS MEG, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, 6 AS TA, dbo.INVO_LST.MEGHk, 
                        dbo.HEAD_LST.DATE_N, dbo.INVO_LST.NUMBER, dbo.TCOD_ANBAR.NAMES AS BEDNAME, dbo.INVO_LST.AVRAGE2, dbo.HEAD_LST.FNUMCO, dbo.INVO_LST.id, 
                        ISNULL(dbo.INVO_LST.MANDAH, N'  ') AS mol
  FROM         dbo.HEAD_LST INNER JOIN
                        dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER INNER JOIN
                        dbo.TCOD_ANBAR ON dbo.HEAD_LST.ANBAR = dbo.TCOD_ANBAR.CODE
  WHERE     (dbo.HEAD_LST.TAG = 5)
        UNION
       SELECT     dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGHk * - 1 AS MEG, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.HEAD_LST.TAG, 
                             dbo.INVO_LST.MEGHk, dbo.HEAD_LST.DATE_N, dbo.INVO_LST.NUMBER, dbo.CUST_HESAB.NAME, dbo.INVO_LST.AVRAGE, dbo.HEAD_LST.FNUMCO, 
                             dbo.INVO_LST.id, ISNULL(dbo.INVO_LST.MANDAH, N'  ') AS mol
       FROM         dbo.HEAD_LST LEFT OUTER JOIN
                             dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes INNER JOIN
                             dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
       WHERE     (dbo.HEAD_LST.TAG = 2) OR
                             (dbo.HEAD_LST.TAG = 8) OR
                             (dbo.HEAD_LST.TAG = 10) OR
                             (dbo.HEAD_LST.TAG = 11) OR
                             (dbo.HEAD_LST.TAG = 26)
        UNION
        SELECT     dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH_MAR * - 1 AS MEG, dbo.INVO_LST.MABL, 
                             dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR AS Expr1, dbo.BACK_HEAD.ta + 2 AS Expr2, dbo.INVO_LST.MEGH_MAR, dbo.BACK_HEAD.DATE_N, 
                             dbo.BACK_HEAD.NUMBER, dbo.CUST_HESAB.NAME, dbo.INVO_LST.AVRAGE2, dbo.BACK_HEAD.FNUMCO, dbo.INVO_LST.id, ISNULL(dbo.INVO_LST.MANDAH, N'  ') AS mol
       
       FROM         dbo.BACK_HEAD INNER JOIN
                             dbo.INVO_LST ON dbo.BACK_HEAD.ta = dbo.INVO_LST.TAG AND dbo.BACK_HEAD.NUMBER1 = dbo.INVO_LST.NUMBER INNER JOIN
                             dbo.CUST_HESAB ON dbo.BACK_HEAD.CUST_NO = dbo.CUST_HESAB.hes
       WHERE     (dbo.BACK_HEAD.ta + 2 = 3) AND (dbo.INVO_LST.MEGH_MAR <> 0)
        UNION
        SELECT     dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH_MAR AS MEG, dbo.INVO_LST.MABL, 
                             dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR AS MABL_K, dbo.BACK_HEAD.ta + 2 AS TAG, dbo.INVO_LST.MEGH_MAR AS MEGHk, 
                             dbo.BACK_HEAD.DATE_N, dbo.BACK_HEAD.NUMBER, dbo.CUST_HESAB.NAME, dbo.INVO_LST.AVRAGE2, dbo.BACK_HEAD.FNUMCO, dbo.INVO_LST.id, ISNULL(dbo.INVO_LST.MANDAH, N'  ') AS mol
       
       FROM         dbo.BACK_HEAD INNER JOIN
                             dbo.INVO_LST ON dbo.BACK_HEAD.ta = dbo.INVO_LST.TAG AND dbo.BACK_HEAD.NUMBER1 = dbo.INVO_LST.NUMBER INNER JOIN
                       dbo.CUST_HESAB ON dbo.BACK_HEAD.CUST_NO = dbo.CUST_HESAB.hes
       WHERE     (dbo.BACK_HEAD.ta + 2 = 4) AND (dbo.INVO_LST.MEGH_MAR <> 0)
        UNION
        SELECT     dbo.ANBGRD_HEAD.GRD_ANBAR, dbo.ANBGRD_LST.CODE, (dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1 AS MEG, 
                             dbo.ANBGRD_LST.MABL, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL AS MABLK, 
                             dbo.UIIF(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3, N'>', 0, 18, 17) AS TAG, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) 
                             AS MEGK, dbo.ANBGRD_HEAD.GRD_DATE, dbo.ANBGRD_LST.GRD_NUM, '            ' AS BEDNAME, dbo.ANBGRD_LST.MABL AS Expr1, 0 AS Expr2, 
                             dbo.ANBGRD_LST.GRD_NUM AS Expr3, ' ' AS mol
       FROM         dbo.ANBGRD_LST INNER JOIN
                             dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM
       WHERE     (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) AND ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1 <> 0)
        union
      SELECT     dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH_MAR AS MEG, dbo.INVO_LST.MABL AS Expr1, 
                         dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR AS mabl_k, dbo.HEAD_LST.TAG, dbo.INVO_LST.MEGH_MAR AS meghk, dbo.HEAD_LST.DATE_N, 
                         dbo.INVO_LST.NUMBER, dbo.CUST_HESAB.NAME, dbo.INVO_LST.AVRAGE, dbo.HEAD_LST.FNUMCO, dbo.INVO_LST.id, 
                         ISNULL(dbo.INVO_LST.MANDAH, N'  ') AS mol
   FROM         dbo.HEAD_LST INNER JOIN
                         dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER LEFT OUTER JOIN
                         dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes
   WHERE     (dbo.HEAD_LST.TAG = 22) AND (dbo.INVO_LST.MEGH_MAR <> 0)
       union
  SELECT     dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGHk * - 1 AS MEG, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.HEAD_LST.TAG, 
                        dbo.INVO_LST.MEGHk, dbo.HEAD_LST.DATE_N, dbo.INVO_LST.NUMBER, dbo.TCOD_ANBAR.NAMES AS BEDNAME, dbo.INVO_LST.AVRAGE, dbo.HEAD_LST.FNUMCO, 
                        dbo.INVO_LST.id, ISNULL(dbo.INVO_LST.MANDAH, N'  ') AS mol
  FROM         dbo.HEAD_LST INNER JOIN
                        dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER INNER JOIN
                        dbo.TCOD_ANBAR ON dbo.HEAD_LST.ANBARF = dbo.TCOD_ANBAR.CODE
  WHERE     (dbo.HEAD_LST.TAG = 5)
 union
     SELECT     dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGHk * - 1 AS MEG, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.HEAD_LST.TAG, 
                             dbo.INVO_LST.MEGHk, dbo.HEAD_LST.DATE_N, dbo.INVO_LST.NUMBER, dbo.CUST_HESAB.NAME, dbo.INVO_LST.AVRAGE, dbo.HEAD_LST.FNUMCO, 
                             dbo.INVO_LST.id, ISNULL(dbo.INVO_LST.MANDAH, N'  ') AS mol
       FROM         dbo.HEAD_LST LEFT OUTER JOIN
                             dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes INNER JOIN
                             dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
       WHERE     (dbo.HEAD_LST.TAG = 20)   AND (dbo.HEAD_LST.tamir=1 OR dbo.HEAD_LST.TAMIR = 4)
   )
   

