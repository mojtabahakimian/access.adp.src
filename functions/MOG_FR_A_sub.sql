CREATE FUNCTION [dbo].[MOG_FR_A_sub]
 (@FORMS___F_MENU_ANBAR_TARAZ___DT2 bigint)
 RETURNS TABLE
 AS
 RETURN ( SELECT     dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS MEG, 
                       SUM(dbo.INVO_LST.AVRAGE * dbo.INVO_LST.MEGHk - ISNULL(dbo.INVO_LST.AVRAGE2, 0) * dbo.INVO_LST.MEGH_MAR) AS avgofmabl
 FROM         dbo.INVO_LST INNER JOIN
                       dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG
 WHERE     (dbo.INVO_LST.TAG = 2 OR
                       dbo.INVO_LST.TAG = 5 OR
                       dbo.INVO_LST.TAG = 8 OR
                       dbo.INVO_LST.TAG = 10 OR
                       dbo.INVO_LST.TAG = 11 OR
                       dbo.INVO_LST.TAG = 26 ) AND (dbo.HEAD_LST.DATE_N <= @FORMS___F_MENU_ANBAR_TARAZ___DT2)
 GROUP BY dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE
 UNION
 SELECT     dbo.ANBGRD_HEAD.GRD_ANBAR AS anbar, dbo.ANBGRD_LST.CODE, dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3 AS MEG, 
                       ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL AS MABLK
 FROM         dbo.ANBGRD_LST INNER JOIN
                       dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM
 WHERE     (dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3 > 0) AND (dbo.ANBGRD_HEAD.GRD_DATE <= @FORMS___F_MENU_ANBAR_TARAZ___DT2) AND 
                       (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) )
