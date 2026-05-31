CREATE FUNCTION [dbo].[Q_GARDESH_KHFR_DAFTAR]
(@Forms___F_MENU_KOL_MOIN_tafzil___DT2 BIGINT,
@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 BIGINT,
@Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF nvarchar(40))
RETURNS TABLE
AS
RETURN ( SELECT     TOP 100 PERCENT SDATE AS SSADTE, N_S, CUST_NO, NAME AS NAM, SHARH, MAS, DATE_N AS SDDSTE, MEGK, MABL, dbo.UIIF(SDATE, N'<=', 
                      @Forms___F_MENU_KOL_MOIN_tafzil___DT2, bed, 0) AS bedd, dbo.UIIF(SDATE, N'<=', @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2, bes, 0) 
                      AS bess, dbo.UIIF(SDATE, N'<=', @Forms___F_MENU_KOL_MOIN_tafzil___DT2, bed, 0) - dbo.UIIF(SDATE, N'<=', 
                      @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2, bes, 0) AS MAND
FROM         dbo.Q_GARDESH_KHFR_DAFTAR_SUB(@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1, @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2, 
                      @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF) Q_GARDESH_KHFR_DAFTAR_SUB
ORDER BY DATE_N )
