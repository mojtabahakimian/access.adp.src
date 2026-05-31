CREATE FUNCTION [dbo].[AVRAGE_VORUD]
(@Forms___PARA___ANBAR int,
@Forms___PARA___DT bigint,
@Forms___PARA___id int,
@Forms___PARA___COD nvarchar(15))
RETURNS TABLE
AS
RETURN ( SELECT     SUM(SumOfMEGHk) AS VMEGHK, SUM(SumOfMABL_K) AS VMABK, CODE
FROM         dbo.AVRAGE_VORUD_SUB(@Forms___PARA___ANBAR, @Forms___PARA___DT, @Forms___PARA___id, @Forms___PARA___COD) AVRAGE_VORUD_SUB
GROUP BY CODE )
