 
   
        CREATE FUNCTION dbo.AK_MOGO_AVL_KOL_SUB
        (@Forms___F_MENU_ANBAR___DT2 bigint,
        @Forms___F_MENU_ANBAR___MANBAR NVARCHAR(10))
        RETURNS TABLE
        AS
        RETURN ( SELECT     CODE, SUM(MOGODI_A) AS MEG, SUM(MABL_A) AS SumOfMABL_A, ANBAR, 0 AS AA
        FROM         dbo.STUF_FSK
        GROUP BY CODE, ANBAR
        HAVING      (ANBAR LIKE @Forms___F_MENU_ANBAR___MANBAR) 
        UNION
        SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS MEG, SUM(dbo.INVO_LST.MABL_K) AS SumOfMABL_K, 
                              dbo.INVO_LST.ANBAR, 1 AS AA
        FROM         dbo.HEAD_LST INNER JOIN
                              dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
        WHERE     (dbo.INVO_LST.TAG = 1 OR
                              dbo.INVO_LST.TAG = 7 OR
                              dbo.INVO_LST.TAG = 9 OR
                              dbo.INVO_LST.TAG =24 ) AND (dbo.HEAD_LST.DATE_N <= @Forms___F_MENU_ANBAR___DT2)
        GROUP BY dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR
        HAVING      (dbo.INVO_LST.ANBAR LIKE @Forms___F_MENU_ANBAR___MANBAR)
        UNION
      SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGH_MAR) AS MEG, SUM(dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR) AS SumOfMABL_K, 
                            dbo.INVO_LST.ANBAR, 1 AS AA
      FROM         dbo.HEAD_LST INNER JOIN
                            dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
      WHERE     (dbo.INVO_LST.TAG = 22) AND (dbo.HEAD_LST.DATE_N <= @Forms___F_MENU_ANBAR___DT2)
      GROUP BY dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR
      HAVING      (dbo.INVO_LST.ANBAR LIKE @Forms___F_MENU_ANBAR___MANBAR)
       union
        SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS MEG, SUM(dbo.INVO_LST.MABL_K) AS SumOfMABL_K, 
                              dbo.INVO_LST.ANBARF, 2 AS AA
        FROM         dbo.HEAD_LST INNER JOIN
                              dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
        WHERE     (dbo.INVO_LST.TAG = 5) AND (dbo.HEAD_LST.DATE_N <= @Forms___F_MENU_ANBAR___DT2)
        GROUP BY dbo.INVO_LST.CODE, dbo.INVO_LST.ANBARF
        HAVING      (dbo.INVO_LST.ANBARF LIKE @Forms___F_MENU_ANBAR___MANBAR)
        union
        SELECT     dbo.ANBGRD_LST.CODE, SUM((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1) AS MEG, 
                           SUM(ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL) AS MABLK, dbo.ANBGRD_HEAD.GRD_ANBAR AS anbar, 
                           3 AS aa
     FROM         dbo.ANBGRD_LST INNER JOIN
                           dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM
     WHERE     (dbo.ANBGRD_HEAD.GRD_DATE <= @Forms___F_MENU_ANBAR___DT2) AND (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) AND 
                           (dbo.ANBGRD_HEAD.GRD_ANBAR LIKE @Forms___F_MENU_ANBAR___MANBAR)
     GROUP BY dbo.ANBGRD_LST.CODE, dbo.ANBGRD_HEAD.GRD_ANBAR
     HAVING      (SUM((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1) >= 0)
  
  )

