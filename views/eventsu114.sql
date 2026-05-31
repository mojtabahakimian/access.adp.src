create view  [dbo].[eventsu114] as  select   TOP (100) PERCENT * from (SELECT TASKS.* FROM TASKS where TASKS.PERSONEL = 114 and TASKS.STATUS = 1) as jafar ORDER BY IDNUM
