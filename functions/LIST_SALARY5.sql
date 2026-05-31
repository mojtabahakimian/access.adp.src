CREATE FUNCTION [dbo].[LIST_SALARY5](@Forms___F_MENU_BIM___MMO INT,
@Forms___Baseknow___PERSONEL INT,
@Forms___Baseknow___PERVAM INT)
RETURNS TABLE
AS
RETURN ( SELECT     dbo.PERSONEL.CODE, dbo.PERSONEL.PNAME + ' ' + dbo.PERSONEL.PFAMILY AS PRF, dbo.PERSONEL.PFAMILY, dbo.PERSONEL.KHNOWNUM, 
                      dbo.PERSONEL.FATHER, dbo.Umonth(dbo.WORKING.WDATE) AS MM, dbo.WORKING.WDATE, dbo.WORKING.DAYS, Q_MAND_SALARY.BEDH, 
                      MAND_VAM.VMAN, dbo.PERSONEL.CODE AS HES_T
FROM         dbo.WORKHEAD INNER JOIN
                      dbo.PERSONEL LEFT OUTER JOIN
                      dbo.Q_MAND_SALARY(@Forms___F_MENU_BIM___MMO, @Forms___Baseknow___PERSONEL) Q_MAND_SALARY ON 
                      dbo.PERSONEL.CODE = Q_MAND_SALARY.HES_T INNER JOIN
                      dbo.WORKING LEFT OUTER JOIN
                      dbo.MAND_VAM(@Forms___Baseknow___PERVAM, @Forms___F_MENU_BIM___MMO) MAND_VAM ON 
                      dbo.WORKING.PCODE = MAND_VAM.HES_T ON dbo.PERSONEL.CODE = dbo.WORKING.PCODE ON 
                      dbo.WORKHEAD.WDATE = dbo.WORKING.WDATE
WHERE     (dbo.Umonth(dbo.WORKING.WDATE) = @Forms___F_MENU_BIM___MMO) )
