create view [dbo].[roundo] as SELECT     BES, BED, ROUND(BED, 0) AS BD, ROUND(BES, 0) AS BS, N_S, HES, SHARH FROM dbo.DEED_DTL WHERE     (ROUND(BED, 0) <> BED) OR (ROUND(BES, 0) <> BES)
