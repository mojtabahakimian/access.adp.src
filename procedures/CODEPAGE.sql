CREATE PROCEDURE [dbo].[CODEPAGE]
 
  AS
 DECLARE @Table NVARCHAR(4000),@Column NVARCHAR(4000)
 
 DECLARE Table_Cursor CURSOR
 
 FOR
 
 --پيدا کردن تمام فيلدهاي متني تمام جداول ديتابيس جاري
 
 SELECT T.name, /* Table */
 
        C.name  /* Column */
 
 FROM sysobjects T,syscolumns C
 
 WHERE  T.id = C.id  AND T.xtype = 'u' /* User Table */
 
          			AND (C.xtype = 99 /* ntext */
 
 					OR C.xtype = 35   /* text */
 
 					OR C.xtype = 231  /* nvarchar */
 
 					OR C.xtype = 167  /* varchar */
 
 					OR C.xtype = 175  /* char */
 
 					OR C.xtype = 239  /* nchar */)
 
 					
 
 OPEN Table_Cursor FETCH NEXT FROM Table_Cursor INTO @Table,@Column
 
 WHILE (@@FETCH_STATUS = 0)
 
 BEGIN
 
 	EXEC ('Update [' + @Table + '] Set [' + @Column + '] = REPLACE(REPLACE(CAST([' + @Column +	'] as nvarchar(4000)), NCHAR(1610), NCHAR(1740)), NCHAR(1603), NCHAR(1705))')
 
 FETCH NEXT FROM Table_Cursor INTO @Table,@Column
 
 END CLOSE Table_Cursor 
 
 DEALLOCATE Table_Cursor
