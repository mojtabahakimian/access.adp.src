CREATE VIEW [dbo].[B_MOG_ANBARHA]
AS
SELECT     dbo.B_MOG_AVL_KOL.ANBAR, dbo.B_MOG_AVL_KOL.CODE, dbo.B_MOG_AVL_KOL.SumOfMEG - ISNULL(dbo.b_mog_fr.MEG, 0) AS MAND, 
                      ISNULL(dbo.b_mog_fr.MEG, 0) AS MEGF, dbo.B_MOG_AVL_KOL.SumOfMEG
FROM         dbo.B_MOG_AVL_KOL LEFT OUTER JOIN
                      dbo.b_mog_fr ON dbo.B_MOG_AVL_KOL.CODE = dbo.b_mog_fr.CODE AND dbo.B_MOG_AVL_KOL.ANBAR = dbo.b_mog_fr.ANBAR
