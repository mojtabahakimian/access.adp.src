create view [dbo].[AGHLAM]
 as
 SELECT     NUMBER, TAG, CODE FROM         dbo.INVO_LST GROUP BY NUMBER, TAG, CODE HAVING      (TAG = 2) OR   (TAG = 14)
