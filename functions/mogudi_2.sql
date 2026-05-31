create FUNCTION [dbo].[mogudi_2]
      (@Forms___F_MENU_ANBAR___MANBAR int)
      RETURNS TABLE
      AS
      RETURN (SELECT     CODE, SUM(MEGHk - MEGH_MAR) AS MEG, ANBAR, 0 AS kk
 FROM         dbo.INVO_LST
 WHERE     (TAG = 2 OR
                       TAG = 5 OR
                       TAG = 8 OR
                       TAG = 10 OR
                       TAG = 11 OR
                       TAG = 26)
 GROUP BY CODE, ANBAR
 HAVING      (ANBAR= @Forms___F_MENU_ANBAR___MANBAR)
      UNION
 SELECT     dbo.ANBGRD_LST.CODE, SUM(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS MEG, dbo.ANBGRD_HEAD.GRD_ANBAR, 1 AS kk
 FROM         dbo.ANBGRD_LST INNER JOIN
                       dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM
 WHERE     (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) AND (dbo.ANBGRD_HEAD.GRD_ANBAR = @Forms___F_MENU_ANBAR___MANBAR)
 GROUP BY dbo.ANBGRD_LST.CODE, dbo.ANBGRD_HEAD.GRD_ANBAR
 HAVING      (SUM(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) > 0)
    UNION
  SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk) AS MEG, dbo.INVO_LST.ANBAR, 4 AS AA
 FROM         dbo.HEAD_LST INNER JOIN
                       dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
 WHERE     (dbo.INVO_LST.TAG = 20) AND (dbo.HEAD_LST.TAMIR = 1)
 GROUP BY dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR
 HAVING      (dbo.INVO_LST.ANBAR = @Forms___F_MENU_ANBAR___MANBAR)
  
  )
