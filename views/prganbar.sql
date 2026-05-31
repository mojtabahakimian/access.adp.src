CREATE VIEW [dbo].[prganbar]
AS
SELECT     dbo.QPROGPAS2.PRGID, dbo.QPROGPAS2.CODB, dbo.MOGUDI_KOL_ANBARHA.NAME, dbo.QPROGPAS2.VAHED, 
                      dbo.QPROGPAS2.SumOfKOLMAV AS kolm, dbo.MOGUDI_KOL_ANBARHA.MANDAH, dbo.STUF_DEF.MIN_M, 
                      dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M AS MOGMOG, 
                      dbo.UIIF(dbo.QPROGPAS2.SumOfKOLMAV * dbo.QPROGPAS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), N'>', 
                      0, dbo.QPROGPAS2.SumOfKOLMAV * dbo.QPROGPAS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), 0) AS sef, 
                      dbo.UIIF(dbo.QPROGPAS2.SumOfKOLMAV * dbo.QPROGPAS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), N'<', 
                      0, (dbo.QPROGPAS2.SumOfKOLMAV * dbo.QPROGPAS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M)) * - 1, 0) 
                      AS EZAF
FROM         dbo.QPROGPAS2 INNER JOIN
                      dbo.MOGUDI_KOL_ANBARHA ON dbo.QPROGPAS2.CODB = dbo.MOGUDI_KOL_ANBARHA.CODE INNER JOIN
                      dbo.STUF_DEF ON dbo.MOGUDI_KOL_ANBARHA.CODE = dbo.STUF_DEF.CODE
