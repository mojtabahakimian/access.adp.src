CREATE FUNCTION [dbo].[LIST_SALARY22_10]
(@Forms___F_MENU_BIM___MMO INT,
@Forms___Baseknow___HOLA INT,
@Forms___Baseknow___HSAY INT,
@Forms___Baseknow___HJAZ INT,
@Forms___Baseknow___HNAH INT,
@Forms___Baseknow___HCON INT,
@Forms___Baseknow___HKHA INT,
@HOKMDATE BIGINT,
@Forms___Baseknow___SAGHFH INT,
@Forms___Baseknow___HSANP INT,
@Forms___Baseknow___HSHI INT,
@Forms___Baseknow___YEA SMALLINT,
@Forms___Baseknow___HEZA INT,
@Forms___Baseknow___HBON INT)
RETURNS TABLE
AS
RETURN ( SELECT     LIST_SALARY_END_10.CODE, LIST_SALARY_END_10.PNAME, LIST_SALARY_END_10.PFAMILY, 
                      LIST_SALARY_END_10.PNAME + ' ' + LIST_SALARY_END_10.PFAMILY AS PNPF, LIST_SALARY_END_10.KHNOWNUM, 
                      LIST_SALARY_END_10.FATHER, LIST_SALARY_END_10.BIMEH_NUM, LIST_SALARY_END_10.JOB, LIST_SALARY_END_10.SEX, 
                      LIST_SALARY_END_10.WSDAT, LIST_SALARY_END_10.WEDATE, LIST_SALARY_END_10.WDATE, LIST_SALARY_END_10.DAYS, 
                      LIST_SALARY_END_10.SALARY_DAYLY, LIST_SALARY_END_10.DAYS * LIST_SALARY_END_10.SALARY_DAYLY AS mosalary, 
                      LIST_SALARY_END_10.HOME * LIST_SALARY_END_10.DAYS AS hom, 
                      LIST_SALARY_END_10.DAYS * LIST_SALARY_END_10.SALARY_DAYLY + dbo.UIIF(LIST_SALARY_END_10.DAYS, N'=', 0, 0, 
                      LIST_SALARY_END_10.MAZMASH) AS jmazsalmash, 
                      LIST_SALARY_END_10.DAYS * LIST_SALARY_END_10.SALARY_DAYLY + dbo.UIIF(LIST_SALARY_END_10.DAYS, N'=', 0, 0, 
                      LIST_SALARY_END_10.MAZMASH + dbo.UIIF(@Forms___Baseknow___HOLA, N'=', - 1, dbo.UIIF(LIST_SALARY_END_10.DAYS, N'>', 27, 
                      LIST_SALARY_END_10.CHILDREN, LIST_SALARY_END_10.CHILDREN / 30 * LIST_SALARY_END_10.DAYS), 0)) AS jamazsalmash, 
                      dbo.UIIF(LIST_SALARY_END_10.DIHARD, N'=', 1, 0, CONVERT(INT, 
                      (LIST_SALARY_END_10.DAYS * LIST_SALARY_END_10.SALARY_DAYLY + dbo.UIIF(LIST_SALARY_END_10.DAYS, N'=', 0, 0, 
                      LIST_SALARY_END_10.MAZMASH)) * 0.07)) AS bimper, 
                      LIST_SALARY_END_10.DAYS * LIST_SALARY_END_10.SALARY_DAYLY + dbo.UIIF(LIST_SALARY_END_10.DAYS, N'=', 0, 0, 
                      LIST_SALARY_END_10.MAZMASH + dbo.UIIF(@Forms___Baseknow___HOLA, N'=', - 1, LIST_SALARY_END_10.CHILDREN, 0)) - CONVERT(INT, 
                      (LIST_SALARY_END_10.DAYS * LIST_SALARY_END_10.SALARY_DAYLY + dbo.UIIF(LIST_SALARY_END_10.DAYS, N'=', 0, 0, 
                      LIST_SALARY_END_10.MAZMASH)) * 0.07) - dbo.PLIST_MALIAT.MALIAT AS ghabel, LIST_SALARY_END_10.CHILDREN AS chid, 
                      LIST_SALARY_END_10.CONDITIONS, LIST_SALARY_END_10.JAZB, LIST_SALARY_END_10.SAYER, LIST_SALARY_END_10.EZAFAH, 
                      LIST_SALARY_END_10.PADASH, LIST_SALARY_END_10.KASR_VAM, dbo.Umonth(LIST_SALARY_END_10.WDATE) AS MM, CONVERT(INT, 
                      (LIST_SALARY_END_10.DAYS * LIST_SALARY_END_10.SALARY_DAYLY + LIST_SALARY_END_10.HOME * LIST_SALARY_END_10.DAYS) * 0.07) 
                      AS bimper2, dbo.UIIF(LIST_SALARY_END_10.TAB56, N'=', 0, 0, 
                      LIST_SALARY_END_10.DAYS * LIST_SALARY_END_10.SALARY_DAYLY + LIST_SALARY_END_10.HOME * LIST_SALARY_END_10.DAYS) 
                      AS JAMMAHS20, LIST_SALARY_END_10.DSW_IDPLC, LIST_SALARY_END_10.DSW_NAT, LIST_SALARY_END_10.DSW_BDATE, 
                      LIST_SALARY_END_10.MaxOfHDATE, LIST_SALARY_END_10.MAZMASH, 
                      CAST(dbo.UIIF(LIST_SALARY_END_10.DAYS * LIST_SALARY_END_10.SALARY_DAYLY + LIST_SALARY_END_10.CHI + dbo.UIIF(LIST_SALARY_END_10.DAYS,
                       N'=', 0, 0, LIST_SALARY_END_10.MAZMASH), N'>', @Forms___Baseknow___SAGHFH, 
                      (LIST_SALARY_END_10.DAYS * LIST_SALARY_END_10.SALARY_DAYLY + LIST_SALARY_END_10.CHI + dbo.UIIF(LIST_SALARY_END_10.DAYS, N'=', 
                      0, 0, LIST_SALARY_END_10.MAZMASH) - @Forms___Baseknow___SAGHFH) * 0.1, 0) AS int) AS mali, LIST_SALARY_END_10.TAB56, 
                      LIST_SALARY_END_10.MELLICOD
FROM         dbo.LIST_SALARY_END_10(@Forms___F_MENU_BIM___MMO, @Forms___Baseknow___HOLA, @Forms___Baseknow___HSAY, 
                      @Forms___Baseknow___HJAZ, @Forms___Baseknow___HNAH, @Forms___Baseknow___HCON, @Forms___Baseknow___HKHA, @HOKMDATE, 
                      @Forms___Baseknow___HSANP, @Forms___Baseknow___HSHI, @Forms___Baseknow___YEA, @Forms___Baseknow___HEZA, 
                      @Forms___Baseknow___HBON) LIST_SALARY_END_10 INNER JOIN
                      dbo.PLIST_MALIAT ON LIST_SALARY_END_10.WDATE = dbo.PLIST_MALIAT.WDATE AND 
                      LIST_SALARY_END_10.CODE = dbo.PLIST_MALIAT.CODE
WHERE     (LIST_SALARY_END_10.BIMEH_NUM > 50) AND (NOT (LIST_SALARY_END_10.BIMEH_NUM IS NULL)) AND 
                      (dbo.Umonth(LIST_SALARY_END_10.WDATE) = @Forms___F_MENU_BIM___MMO) )
