create view  [dbo].[eventsu87] as  select   TOP (100) PERCENT * from (SELECT TASKS.* FROM TASKS where TASKS.PERSONEL = 87 and TASKS.STATUS = 1) as jafar ORDER BY IDNUM
