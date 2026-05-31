CREATE VIEW [dbo].[SALS]
AS
SELECT     dbo.TFORMS.CAPTION, dbo.TFORMS.kind, dbo.SAL_CHEK.USERCO, dbo.SAL_CHEK.OBJECT, dbo.SAL_CHEK.RUN, dbo.SAL_CHEK.SEE, 
                      dbo.SAL_CHEK.INP, dbo.SAL_CHEK.UPD, dbo.SAL_CHEK.DEL, dbo.SALA_DTL.SAL_NAME
FROM         dbo.TFORMS INNER JOIN
                      dbo.SAL_CHEK ON dbo.TFORMS.IDH = dbo.SAL_CHEK.OBJECT INNER JOIN
                      dbo.SALA_DTL ON dbo.SAL_CHEK.USERCO = dbo.SALA_DTL.IDD
