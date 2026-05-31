CREATE FUNCTION [dbo].[LIST_SALARY22TOFILE_10]
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
 @Forms___Baseknow___HSHI int,
 @Forms___Baseknow___YEA SMALLINT,
 @Forms___Baseknow___HEZA INT,
 @Forms___Baseknow___HBON  INT)
 RETURNS TABLE
 AS
 RETURN ( SELECT     COUNT(CODE) AS DSK_NUM, SUM(DAYS) AS DSK_TDD, SUM(dbo.UIIF(DAYS, N'>', 0, SALARY_DAYLY, 0)) AS DSK_TROOZ, SUM(mosalary) 
                       AS DSK_TMAH, SUM(MAZMASH) AS DSK_TMAZ, SUM(jamazsalmash) AS DSK_TTOTL, SUM(bimper) AS DSK_TBIME, SUM(dbo.UIIF(TAB56, N'=', 0, 
                       jmazsalmash, 0) * 0.2) AS DSK_TKOSO, SUM(jmazsalmash * 0.03) AS DSK_BIC, WDATE, SUM(jmazsalmash) AS DSK_TMASH
 FROM         dbo.LIST_SALARY22_10(@Forms___F_MENU_BIM___MMO, @Forms___Baseknow___HOLA, @Forms___Baseknow___HSAY, 
                       @Forms___Baseknow___HJAZ, @Forms___Baseknow___HNAH, @Forms___Baseknow___HCON, @Forms___Baseknow___HKHA, @HOKMDATE, 
                       @Forms___Baseknow___SAGHFH, @Forms___Baseknow___HSANP, @Forms___Baseknow___HSHI, @Forms___Baseknow___YEA, 
                       @Forms___Baseknow___HEZA, @Forms___Baseknow___HBON ) LIST_SALARY22_10
 GROUP BY WDATE )
