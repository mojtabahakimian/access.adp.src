CREATE FUNCTION [dbo].[AVRAGE_MOG](@Forms___PARA___ANBAR int,
@Forms___PARA___DT bigint,
@Forms___PARA___id int,
@Forms___PARA___COD nvarchar(15))
RETURNS TABLE
AS
RETURN ( SELECT     dbo.STUF_DEF.CODE, AVRAGE_VORUD.VMEGHK - ISNULL(AVREG_MOG_FR.FMEGHK, 0) AS MOG, 
                      ABS(AVRAGE_VORUD.VMABK - ISNULL(AVREG_MOG_FR.FMABK, 0)) AS mabl, AVRAGE_VORUD.VMEGHK, AVRAGE_VORUD.VMABK, 
                      AVREG_MOG_FR.FMABK, AVREG_MOG_FR.FMEGHK
FROM         dbo.AVRAGE_VORUD(@Forms___PARA___ANBAR, @Forms___PARA___DT, @Forms___PARA___id, @Forms___PARA___COD) 
                      AVRAGE_VORUD RIGHT OUTER JOIN
                      dbo.STUF_DEF ON AVRAGE_VORUD.CODE = dbo.STUF_DEF.CODE LEFT OUTER JOIN
                      dbo.AVREG_MOG_FR(@Forms___PARA___ANBAR, @Forms___PARA___DT, @Forms___PARA___id, @Forms___PARA___COD) AVREG_MOG_FR ON 
                      dbo.STUF_DEF.CODE = AVREG_MOG_FR.CODE
WHERE     (dbo.STUF_DEF.CODE = @Forms___PARA___COD) )
