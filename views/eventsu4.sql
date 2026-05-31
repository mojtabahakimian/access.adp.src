create view  [dbo].[eventsu4] as  select   TOP (100) PERCENT * from (SELECT TASKS.* FROM TASKS where TASKS.PERSONEL = 4 and TASKS.STATUS = 1) as jafar ORDER BY IDNUM
