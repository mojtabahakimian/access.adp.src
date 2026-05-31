CREATE FUNCTION [dbo].[LIST_SALARY_END_10]
(@Forms___F_MENU_BIM___MMO INT,
@Forms___Baseknow___HOLA INT,
@Forms___Baseknow___HSAY INT,
@Forms___Baseknow___HJAZ INT,
@Forms___Baseknow___HNAH INT,
@Forms___Baseknow___HCON INT,
@Forms___Baseknow___HKHA INT,
@HOKMDATE BIGINT,
@Forms___Baseknow___HSANP INT,
@Forms___Baseknow___HSHI INT,
@Forms___Baseknow___YEA SMALLINT,
@Forms___Baseknow___HEZA INT,
@Forms___Baseknow___HBON INT)
RETURNS TABLE
AS
RETURN ( SELECT     dbo.PERSONEL.CODE, dbo.PERSONEL.PNAME, dbo.PERSONEL.PFAMILY, dbo.PERSONEL.PNAME + ' ' + dbo.PERSONEL.PFAMILY AS PNPF, 
                      dbo.PERSONEL.KHNOWNUM, dbo.PERSONEL.FATHER, dbo.PERSONEL.BIMEH_NUM, dbo.PERSONEL.JOB, dbo.PERSONEL.SEX, 
                      dbo.UIIF(dbo.Uyear(dbo.PERSONEL.WSDATE) * 100 + dbo.Umonth(dbo.PERSONEL.WSDATE), N'>=', 
                      @Forms___Baseknow___YEA * 100 + @Forms___F_MENU_BIM___MMO, dbo.PERSONEL.WSDATE, NULL) AS WSDAT, dbo.PERSONEL.WEDATE, 
                      dbo.WORKING.WDATE, dbo.WORKING.DAYS, dbo.WORKING.EZAFAH, dbo.WORKING.PADASH, dbo.WORKING.KASR_VAM, 
                      MONTH(dbo.WORKING.WDATE) AS MM, dbo.PERSONEL.DSW_IDPLC, dbo.PERSONEL.DSW_NAT, dbo.PERSONEL.DSW_BDATE, dbo.PHOKM.HONUM, 
                      dbo.PHOKM.HDATE AS MaxOfHDATE, dbo.PHOKM.SALARY_DAYLY + dbo.UIIF(dbo.Umonth(dbo.WORKING.WDATE) 
                      * 100 + dbo.Uday(dbo.WORKING.WDATE), N'>', dbo.Umonth(dbo.PERSONEL.WSDATE) * 100 + dbo.Uday(dbo.PERSONEL.WSDATE), 
                      @Forms___Baseknow___HSANP, 0) AS SALARY_DAYLY, dbo.PHOKM.SALARY_DAYLYB + dbo.UIIF(dbo.Umonth(dbo.WORKING.WDATE) 
                      * 100 + dbo.Uday(dbo.WORKING.WDATE), N'>', dbo.Umonth(dbo.PERSONEL.WSDATE) * 100 + dbo.Uday(dbo.PERSONEL.WSDATE), 
                      @Forms___Baseknow___HSANP, 0) AS SALARY_DAYLYB, dbo.PHOKM.HOME, dbo.PHOKM.CHILDREN, dbo.PHOKM.CONDITIONS, dbo.PHOKM.BON, 
                      dbo.PHOKM.JAZB, dbo.PHOKM.SAYER, dbo.PERSONEL.TAB56, (dbo.UIIF(@Forms___Baseknow___HSAY, N'=', - 1, dbo.PHOKM.SAYER, 0) 
                      + dbo.UIIF(@Forms___Baseknow___HSHI, N'=', - 1, CAST(dbo.PHOKM.HSHIFT / 100 * dbo.PHOKM.SALARY_DAYLY AS INT), 0) 
                      + dbo.UIIF(@Forms___Baseknow___HJAZ, N'=', - 1, dbo.PHOKM.JAZB, 0) + dbo.UIIF(@Forms___Baseknow___HNAH, N'=', - 1, dbo.PHOKM.BON, 0) 
                      + dbo.UIIF(@Forms___Baseknow___HCON, N'=', - 1, dbo.PHOKM.CONDITIONS, 0)) * dbo.WORKING.DAYS + dbo.UIIF(@Forms___Baseknow___HKHA, 
                      N'=', - 1, dbo.UIIF(dbo.WORKING.DAYS, N'>', 27, dbo.PHOKM.HOME, dbo.PHOKM.HOME / 30 * dbo.WORKING.DAYS), 0) 
                      + dbo.UIIF(@Forms___Baseknow___HEZA, N'=', - 1, dbo.WORKING.EZAFAH, 0) + dbo.UIIF(@Forms___Baseknow___HBON, N'=', - 1, 
                      dbo.UIIF(dbo.WORKING.DAYS, N'>', 27, dbo.PHOKM.bonj, dbo.PHOKM.bonj / 30 * dbo.WORKING.DAYS), 0) AS MAZMASH, 
                      dbo.UIIF(@Forms___Baseknow___HOLA, N'=', - 1, dbo.UIIF(dbo.WORKING.DAYS, N'>', 27, dbo.PHOKM.CHILDREN, 
                      dbo.PHOKM.CHILDREN / 30 * dbo.WORKING.DAYS), 0) AS CHI, dbo.UIIF(dbo.Umonth(dbo.WORKING.WDATE) 
                      * 100 + dbo.Uday(dbo.WORKING.WDATE), N'>', dbo.Umonth(dbo.PERSONEL.WSDATE) * 100 + dbo.Uday(dbo.PERSONEL.WSDATE), 
                      @Forms___Baseknow___HSANP, 0) AS Expr1, dbo.PERSONEL.DIHARD, dbo.PERSONEL.MELLICOD
FROM         dbo.PERSONEL INNER JOIN
                      dbo.PHOKM ON dbo.PERSONEL.CODE = dbo.PHOKM.CODE INNER JOIN
                      dbo.WORKING ON dbo.PERSONEL.CODE = dbo.WORKING.PCODE INNER JOIN
                      dbo.LASTHOKM(@HOKMDATE) LASTHOKM ON dbo.PHOKM.HONUM = LASTHOKM.MaxOfHONUM AND dbo.PHOKM.CODE = LASTHOKM.CODE
WHERE     (dbo.Umonth(dbo.WORKING.WDATE) = @Forms___F_MENU_BIM___MMO) )
