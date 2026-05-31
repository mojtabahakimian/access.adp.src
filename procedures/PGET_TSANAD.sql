CREATE PROCEDURE [dbo].[PGET_TSANAD](@Forms___PGET_HED___N_S nvarchar(50),
@Forms___PGET_HED___ID int)
AS INSERT INTO dbo.DEED_DTL
                      (HES_K, HES_M, HES_T, SHARH, BED, N_SERI, BANK, N_S, HES)
SELECT     THES_K, THES_M, THES_T, SHARH, MABL, N_SERI, BANK, @Forms___PGET_HED___N_S AS Expr1, RTRIM(CAST(THES_K AS nvarchar)) 
                      + '-' + RTRIM(CAST(THES_M AS nvarchar)) + '-' + RTRIM(CAST(THES_T AS nvarchar)) AS Expr3
FROM         dbo.PGET_LST
WHERE     (ID = @Forms___PGET_HED___ID)
