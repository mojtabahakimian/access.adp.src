create view  [dbo].[eventscount] as   SELECT     idc, COUNT(idc) AS idcn FROM dbo.CRMEVENTS GROUP BY idc
