CREATE FUNCTION [dbo].[PGET_HED_REP_ASHKHAS]
(@Forms___Baseknow___BESTANKAR int,
@Forms___Baseknow___BEDEHKAR int,
@ID INT)
RETURNS TABLE
AS
RETURN ( SELECT     DT, MOLAH, N_S, FHES_K, FHES_M, FHES_T, KK, FHES, TTNAME, SUM(mabbeD) AS mabbeDD, SUM(mabbeS) AS mabbeSS, SUM(mabbeD - mabbeS)
                       AS mabmbed
FROM         dbo.PGET_HED_REP_ASHKHAS1(@Forms___Baseknow___BESTANKAR, @Forms___Baseknow___BEDEHKAR, @ID) 
                      PGET_HED_REP_ASHKHAS1
GROUP BY DT, MOLAH, N_S, FHES_K, FHES_M, FHES_T, KK, FHES, TTNAME )
