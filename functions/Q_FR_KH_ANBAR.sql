CREATE FUNCTION [dbo].[Q_FR_KH_ANBAR](@Forms___F_MENU_ANBAR_FRKH___DT1 bigint,
@Forms___F_MENU_ANBAR_FRKH___DT2 bigint,
@Forms___F_MENU_ANBAR_FRKH___DTT float,
@Forms___F_MENU_ANBAR_FRKH___MANBAR nvarchar(20))
RETURNS TABLE
AS
RETURN ( SELECT     dbo.HEAD_LST.DATE_N, dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, RTRIM(CAST(dbo.INVO_LST.ANBAR AS NVARCHAR)) AS ANBAR, 
                      dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.VAHED_K, 
                      dbo.TCOD_VAHEDS.NAMES, dbo.TCOD_ANBAR.NAMES AS TA_NAMES, dbo.STUF_DEF.NAME
FROM         dbo.TCOD_VAHEDS INNER JOIN
                      dbo.TCOD_ANBAR INNER JOIN
                      dbo.INVO_LST INNER JOIN
                      dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG ON 
                      dbo.TCOD_ANBAR.CODE = dbo.INVO_LST.ANBAR INNER JOIN
                      dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE ON dbo.TCOD_VAHEDS.CODE = dbo.INVO_LST.VAHED_K
WHERE     (dbo.HEAD_LST.DATE_N BETWEEN @Forms___F_MENU_ANBAR_FRKH___DT1 AND @Forms___F_MENU_ANBAR_FRKH___DT2) AND 
                      (dbo.INVO_LST.TAG = @Forms___F_MENU_ANBAR_FRKH___DTT) AND (RTRIM(CAST(dbo.INVO_LST.ANBAR AS NVARCHAR)) 
                      LIKE @Forms___F_MENU_ANBAR_FRKH___MANBAR) )
