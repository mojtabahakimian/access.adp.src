create FUNCTION [dbo].[mogudi_1]
         (  @Forms___F_MENU_ANBAR___MANBAR int)
         RETURNS TABLE
         AS
         RETURN ( SELECT     CODE, SUM(MOGODI_A) AS MEG, ANBAR, 0 AS AA
 FROM         dbo.STUF_FSK
 GROUP BY CODE, ANBAR
 HAVING      (ANBAR =  @Forms___F_MENU_ANBAR___MANBAR)
 
  UNION
 SELECT     CODE, SUM(MEGHk - MEGH_MAR) AS MEG, ANBAR, 1 AS AA
 FROM         dbo.INVO_LST
 WHERE     (TAG = 1) OR
                       (TAG = 7) OR
                       (TAG = 9) OR
                       (TAG = 24)
 GROUP BY CODE, ANBAR
 HAVING      (ANBAR = @Forms___F_MENU_ANBAR___MANBAR)
         UNION
 SELECT     CODE, SUM(MEGH_MAR) AS MEG, ANBAR, 1 AS AA
 FROM         dbo.INVO_LST
 WHERE     (TAG = 22)
 GROUP BY CODE, ANBAR
 HAVING      (ANBAR = @Forms___F_MENU_ANBAR___MANBAR)
 union
 SELECT     CODE, SUM(MEGHk - MEGH_MAR) AS MEG, ANBARF, 2 AS AA
 FROM         dbo.INVO_LST
 WHERE     (TAG = 5)
 GROUP BY CODE, ANBARF
 HAVING      (ANBARF = @Forms___F_MENU_ANBAR___MANBAR)
         union
 SELECT     dbo.ANBGRD_LST.CODE, SUM((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1) AS MEG, dbo.ANBGRD_HEAD.GRD_ANBAR AS anbar, 3 AS AA
 FROM         dbo.ANBGRD_LST INNER JOIN
                       dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM
 WHERE     (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) AND (dbo.ANBGRD_HEAD.GRD_ANBAR =  @Forms___F_MENU_ANBAR___MANBAR)
 GROUP BY dbo.ANBGRD_LST.CODE, dbo.ANBGRD_HEAD.GRD_ANBAR
 HAVING      (SUM((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1) >= 0)  
   )
