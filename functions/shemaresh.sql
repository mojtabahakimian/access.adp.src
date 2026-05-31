create  FUNCTION [dbo].[shemaresh]
 ( @DT2 bigint)
 RETURNS TABLE
 AS
 RETURN (SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG
FROM         dbo.HEAD_LST INNER JOIN
                      dbo.head_lst_log ON dbo.HEAD_LST.NUMBER = dbo.head_lst_log.NUMBER AND dbo.HEAD_LST.TAG = dbo.head_lst_log.TAGG
GROUP BY dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.TAG
HAVING      (dbo.HEAD_LST.TAG <> 20) AND (MIN(dbo.head_lst_log.UDATEF) = @DT2) )
