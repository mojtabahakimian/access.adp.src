CREATE VIEW [dbo].[zzz_list_mavad_dar_folmul]
AS
SELECT     TOP (100) PERCENT D.CODE AS MaterialCode, dbo.STUF_DEF.NAME
FROM         dbo.DTL_MANF AS D INNER JOIN
                      dbo.HEAD_MANF AS H ON H.FNUMB = D.FNUMB INNER JOIN
                      dbo.STUF_DEF ON D.CODE = dbo.STUF_DEF.CODE
WHERE     (H.CODE IN (2537, 2534, 3258, 10, 2660, 3330, 3220, 2587, 2889)) AND (H.GHEYMAT = 5)
GROUP BY D.CODE, dbo.STUF_DEF.NAME
ORDER BY dbo.STUF_DEF.NAME, MaterialCode
