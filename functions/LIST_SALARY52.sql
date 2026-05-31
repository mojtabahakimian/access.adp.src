CREATE FUNCTION [dbo].[LIST_SALARY52]
 (@Forms___F_MENU_BIM___MMO INT,
 @Forms___Baseknow___PERSONEL INT,
 @Forms___Baseknow___PERVAM INT,
 @MOIN INT)
 RETURNS TABLE
 AS
 RETURN ( SELECT     dbo.PERSONEL.CODE, dbo.PERSONEL.PNAME + ' ' + dbo.PERSONEL.PFAMILY AS PRF, dbo.PERSONEL.PFAMILY, dbo.PERSONEL.KHNOWNUM, 
                       dbo.PERSONEL.FATHER, dbo.Umonth(dbo.WORKING.WDATE) AS MM, dbo.WORKING.WDATE, dbo.WORKING.DAYS, Q_MAND_SALARY2.BEDH, 
                       MAND_VAM2.VMAN, dbo.PERSONEL.CODE AS HES_T
 FROM         dbo.WORKHEAD INNER JOIN
                       dbo.PERSONEL INNER JOIN
                       dbo.WORKING ON dbo.PERSONEL.CODE = dbo.WORKING.PCODE ON dbo.WORKHEAD.WDATE = dbo.WORKING.WDATE LEFT OUTER JOIN
                       dbo.MAND_VAM2(@Forms___Baseknow___PERVAM, @Forms___F_MENU_BIM___MMO, @MOIN) MAND_VAM2 ON 
                       dbo.WORKING.PCODE = MAND_VAM2.HES_T LEFT OUTER JOIN
                       dbo.Q_MAND_SALARY2(@Forms___F_MENU_BIM___MMO, @Forms___Baseknow___PERSONEL, @MOIN) Q_MAND_SALARY2 ON 
                       dbo.PERSONEL.CODE = Q_MAND_SALARY2.HES_T
 WHERE     (dbo.Umonth(dbo.WORKING.WDATE) = @Forms___F_MENU_BIM___MMO) )
