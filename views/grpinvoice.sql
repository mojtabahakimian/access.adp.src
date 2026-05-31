CREATE view  [dbo].[grpinvoice]
as
SELECT     TOP 100 PERCENT dbo.KALAS.DATE_N, dbo.KALAS.code, dbo.KALAS.kala, dbo.KALAS.MEGH, dbo.KALAS.MEGHk, dbo.KALAS.CUSTNAME, dbo.KALAS.hes, 
                      dbo.KALAS.MABL, dbo.KALAS.MABL_K, dbo.KALAS.KHFR, dbo.KALAS.GHFR, dbo.KALAS.VAHCODE, dbo.KALAS.GRPCODE, dbo.KALAS.MOLAH, dbo.KALAS.SHARAYET, 
                      dbo.KALAS.FNUMCO, dbo.KALAS.MANDAH, dbo.KALAS.ANBARCODE, dbo.KALAS.N_S, dbo.KALAS.USER_NAME, dbo.KALAS.MAS, dbo.KALAS.NUMBER, 
                      dbo.KALAS.NUMBER1, dbo.KALAS.BARGAH, dbo.KALAS.TAGCODE, dbo.OTHER_DTL_SUB.CAM_KHALY, dbo.OTHER_DTL_SUB.CAM_POOR, 
                      dbo.OTHER_DTL_SUB.MEGHk AS MEGHkg, dbo.OTHER_DTL_SUB.VAZNH, dbo.OTHER_DTL.CAMIUN_NUM, dbo.OTHER_DTL.CAMIUN, dbo.OTHER_DTL_SUB.TOZIH, 
                      dbo.OTHER_DTL_SUB.VAZNH - dbo.KALAS.MEGHk AS evazn, dbo.KALAS.ANBARAS, dbo.OTHER_DTL.REQUEST_NO, 
                      dbo.OTHER_DTL_SUB.VAZNH * dbo.KALAS.MABL AS mabkv, dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, CAST(dbo.OTHER_DTL_SUB.TOZIH AS int) 
                      AS kerayah, dbo.OTHER_DTL.BARNAMEH
FROM         dbo.KALAS INNER JOIN
                      dbo.OTHER_DTL_SUB ON dbo.KALAS.code = dbo.OTHER_DTL_SUB.CODE AND dbo.KALAS.NUMBER1 = dbo.OTHER_DTL_SUB.NUMBER AND 
                      dbo.KALAS.TAG = dbo.OTHER_DTL_SUB.TAGG + 11 INNER JOIN
                      dbo.OTHER_DTL ON dbo.KALAS.TAG = dbo.OTHER_DTL.TAG + 11 AND dbo.KALAS.NUMBER1 = dbo.OTHER_DTL.NUMBER INNER JOIN
                      dbo.CUST_HESAB ON dbo.KALAS.hes = dbo.CUST_HESAB.hes
WHERE     (dbo.KALAS.TAGCODE = 12)
ORDER BY dbo.KALAS.hes, dbo.KALAS.FNUMCO, dbo.KALAS.NUMBER
