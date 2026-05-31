create view  [dbo].[eventsu112] as  select   TOP (100) PERCENT * from (SELECT TASKS.* FROM TASKS where TASKS.PERSONEL = 112 and TASKS.STATUS = 1) as jafar ORDER BY IDNUM
