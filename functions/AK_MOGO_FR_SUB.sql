 
  
   
   
   
    
     CREATE FUNCTION dbo.AK_MOGO_FR_SUB
     (@Forms___F_MENU_ANBAR___DT2 bigint,
     @Forms___F_MENU_ANBAR___MANBAR NVARCHAR(10))
     RETURNS TABLE
     AS
     RETURN ( SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS MEG, AVG(dbo.INVO_LST.MABL) AS AvgOfMABL, 
                           dbo.INVO_LST.ANBAR,0 as kk
     FROM         dbo.HEAD_LST INNER JOIN
                           dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
     WHERE     (dbo.INVO_LST.TAG = 2 OR
                           dbo.INVO_LST.TAG = 5 OR
                           dbo.INVO_LST.TAG = 8 OR
                           dbo.INVO_LST.TAG = 10 OR
                           dbo.INVO_LST.TAG = 11 OR
                           dbo.INVO_LST.TAG = 26) AND (dbo.HEAD_LST.DATE_N <= @Forms___F_MENU_ANBAR___DT2)
     GROUP BY dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR
     HAVING      (dbo.INVO_LST.ANBAR LIKE @Forms___F_MENU_ANBAR___MANBAR)
     UNION
     SELECT     dbo.ANBGRD_LST.CODE, SUM(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS MEG, 
                          SUM(ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL) AS MABLK, 
                          dbo.ANBGRD_HEAD.GRD_ANBAR AS anbar,1 as kk
    FROM         dbo.ANBGRD_LST INNER JOIN
                          dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM
    WHERE     (dbo.ANBGRD_HEAD.GRD_DATE <= @Forms___F_MENU_ANBAR___DT2) AND (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) AND 
                          (dbo.ANBGRD_HEAD.GRD_ANBAR LIKE  @Forms___F_MENU_ANBAR___MANBAR)
    GROUP BY dbo.ANBGRD_LST.CODE, dbo.ANBGRD_HEAD.GRD_ANBAR
    HAVING      (SUM(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) > 0)
   UNION
     SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHK) AS MEG,  SUM(dbo.INVO_LST.MABL_K) AS SumOfMABL_K, 
                           dbo.INVO_LST.ANBAR, 4 AS AA
     FROM         dbo.HEAD_LST INNER JOIN
                           dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
     WHERE     (dbo.INVO_LST.TAG = 20) AND (dbo.HEAD_LST.DATE_N <= @Forms___F_MENU_ANBAR___DT2)  AND (dbo.HEAD_LST.tamir=1 OR dbo.HEAD_LST.TAMIR = 4)
     GROUP BY dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR
     HAVING      (dbo.INVO_LST.ANBAR LIKE @Forms___F_MENU_ANBAR___MANBAR)
 
 )
 

