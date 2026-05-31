-- Fixed version for SQL Server 2008R2
CREATE PROCEDURE dbo.usp_Audit_GenerateTrigger
  @Schema SYSNAME,
  @Table  SYSNAME,
  @PKList NVARCHAR(400)    -- مثال: 'IDD' یا 'NUMBER,BRANCH'
AS
BEGIN
  SET NOCOUNT ON;

  DECLARE @tblId INT;
  SET @tblId = OBJECT_ID(QUOTENAME(@Schema)+'.'+QUOTENAME(@Table));
  IF @tblId IS NULL
  BEGIN
    RAISERROR(N'جدول یافت نشد.',16,1); RETURN;
  END;

  -- Create temporary table for columns
  CREATE TABLE #TempCols (
    ColName SYSNAME,
    column_id INT
  );

  -- Insert eligible columns
  INSERT INTO #TempCols (ColName, column_id)
  SELECT 
    c.name AS ColName,
    c.column_id
  FROM sys.columns c
  JOIN sys.types t ON t.user_type_id = c.user_type_id
  WHERE c.object_id = @tblId
    AND c.is_computed = 0
    AND t.name NOT IN ('image','text','ntext','sql_variant','xml','geometry','geography','hierarchyid')
    AND NOT (t.name IN ('varbinary') AND c.max_length = -1) -- varbinary(max)
    AND c.is_rowguidcol = 0;

  -- Create temporary table for PKs
  CREATE TABLE #TempPKs (
    Col SYSNAME
  );

  -- Parse PK list and insert into temp table
  -- Using a simple loop instead of table-valued function
  DECLARE @PKTemp NVARCHAR(400);
  DECLARE @Pos INT;
  DECLARE @PKCol SYSNAME;
  
  SET @PKTemp = REPLACE(@PKList,' ','') + ',';
  
  WHILE CHARINDEX(',', @PKTemp) > 0
  BEGIN
    SET @Pos = CHARINDEX(',', @PKTemp);
    SET @PKCol = LTRIM(RTRIM(LEFT(@PKTemp, @Pos - 1)));
    IF LEN(@PKCol) > 0
    BEGIN
      INSERT INTO #TempPKs (Col) VALUES (@PKCol);
    END
    SET @PKTemp = SUBSTRING(@PKTemp, @Pos + 1, LEN(@PKTemp));
  END;

  -- ساخت لیست VALUES برای i./d. (دینامیک)
  DECLARE @cols NVARCHAR(MAX);
  SET @cols = N'';
  
  SELECT @cols = @cols + 
    N' (N''' + REPLACE(ColName,'''','''''') + ''', CONVERT(NVARCHAR(4000), i.'+QUOTENAME(ColName)+'), CONVERT(NVARCHAR(4000), d.'+QUOTENAME(ColName)+')),' + CHAR(10)
  FROM #TempCols
  ORDER BY column_id;
  
  IF LEN(@cols) > 0
     SET @cols = STUFF(@cols, LEN(@cols), 1, ''); -- remove last comma

  -- سازنده‌ی رشته PK: col1=val1|col2=val2 ...
  DECLARE @pkExpr_I NVARCHAR(MAX), @pkExpr_D NVARCHAR(MAX);
  SET @pkExpr_I = N'';
  SET @pkExpr_D = N'';

  SELECT @pkExpr_I = @pkExpr_I + 
         CASE WHEN @pkExpr_I=N'' THEN N'' ELSE N' + N''|'' + ' END +
         N'N''' + p.Col + N'=''+ISNULL(CONVERT(NVARCHAR(4000), i.'+QUOTENAME(p.Col)+N'),N''NULL'')'
  FROM #TempPKs p;

  SELECT @pkExpr_D = @pkExpr_D + 
         CASE WHEN @pkExpr_D=N'' THEN N'' ELSE N' + N''|'' + ' END +
         N'N''' + p.Col + N'=''+ISNULL(CONVERT(NVARCHAR(4000), d.'+QUOTENAME(p.Col)+N'),N''NULL'')'
  FROM #TempPKs p;

  DECLARE @trgName SYSNAME, @drop NVARCHAR(MAX), @joinOn NVARCHAR(MAX);
  SET @trgName = N'trg_Audit_' + @Table;

  -- Drop قدیمی
  SET @drop = N'IF OBJECT_ID(''' + QUOTENAME(@Schema)+'.'+QUOTENAME(@trgName) + N''',''TR'') IS NOT NULL DROP TRIGGER ' 
               + QUOTENAME(@Schema)+'.'+QUOTENAME(@trgName) + N';';
  EXEC (@drop);

  -- ساخت بخش JOIN بر اساس PK
  SET @joinOn = N'';
  SELECT @joinOn = @joinOn + CASE WHEN @joinOn=N'' THEN N'' ELSE N' AND ' END
                    + N'i.'+QUOTENAME(p.Col)+N' = d.'+QUOTENAME(p.Col)
  FROM #TempPKs p;

  -- بدنه‌ی تریگر
  DECLARE @trg NVARCHAR(MAX);
  SET @trg = N'
CREATE TRIGGER ' + QUOTENAME(@Schema)+'.'+QUOTENAME(@trgName) + N'
ON ' + QUOTENAME(@Schema)+'.'+QUOTENAME(@Table) + N'
AFTER INSERT, UPDATE, DELETE
AS
BEGIN
  SET NOCOUNT ON;

  DECLARE @Schema SYSNAME = N' + QUOTENAME(@Schema,'''') + N';
  DECLARE @Table  SYSNAME = N' + QUOTENAME(@Table ,'''') + N';
  DECLARE @ctx NVARCHAR(128) = CONVERT(NVARCHAR(128), CONTEXT_INFO());

  -- INSERT
  IF EXISTS(SELECT 1 FROM inserted) AND NOT EXISTS(SELECT 1 FROM deleted)
  BEGIN
    INSERT dbo.AuditTrail (EventTimeUtc, SchemaName, TableName, PK, Operation, ColumnName, OldValue, NewValue,
                           SqlLogin, HostName, AppName, AppContext, Spid, RowHash, ChainHash)
    SELECT
      SYSUTCDATETIME(), @Schema, @Table,
      ' + @pkExpr_I + N' AS PK,
      ''I'' AS Operation,
      V.ColName,
      NULL AS OldValue,
      V.NewVal,
      ORIGINAL_LOGIN(), HOST_NAME(), APP_NAME(), @ctx, @@SPID,
      HASHBYTES(''SHA1'', CONVERT(VARBINARY(MAX),
         ISNULL(V.ColName,'''')+ISNULL(V.NewVal,'''')+ISNULL(@ctx,'''')
      )),
      HASHBYTES(''SHA1'',
         ISNULL((SELECT TOP(1) ChainHash FROM dbo.AuditTrail
                 WHERE TableName=@Table AND PK = (' + @pkExpr_I + N')
                 ORDER BY AuditId DESC), 0x) +
         CONVERT(VARBINARY(MAX),
           ISNULL(V.ColName,'''')+ISNULL(V.NewVal,'''')+ISNULL(@ctx,'''')
         )
      )
    FROM inserted i
    CROSS APPLY (VALUES
' + @cols + N'
    ) V(ColName, NewVal, OldVal);
  END

  -- DELETE
  IF EXISTS(SELECT 1 FROM deleted) AND NOT EXISTS(SELECT 1 FROM inserted)
  BEGIN
    INSERT dbo.AuditTrail (EventTimeUtc, SchemaName, TableName, PK, Operation, ColumnName, OldValue, NewValue,
                           SqlLogin, HostName, AppName, AppContext, Spid, RowHash, ChainHash)
    SELECT
      SYSUTCDATETIME(), @Schema, @Table,
      ' + @pkExpr_D + N' AS PK,
      ''D'' AS Operation,
      V.ColName,
      V.OldVal,
      NULL AS NewValue,
      ORIGINAL_LOGIN(), HOST_NAME(), APP_NAME(), @ctx, @@SPID,
      HASHBYTES(''SHA1'', CONVERT(VARBINARY(MAX),
         ISNULL(V.ColName,'''')+ISNULL(V.OldVal,'''')+ISNULL(@ctx,'''')
      )),
      HASHBYTES(''SHA1'',
         ISNULL((SELECT TOP(1) ChainHash FROM dbo.AuditTrail
                 WHERE TableName=@Table AND PK = (' + @pkExpr_D + N')
                 ORDER BY AuditId DESC), 0x) +
         CONVERT(VARBINARY(MAX),
           ISNULL(V.ColName,'''')+ISNULL(V.OldVal,'''')+ISNULL(@ctx,'''')
         )
      )
    FROM deleted d
    CROSS APPLY (VALUES
' + REPLACE(@cols,'i.','d.') + N'
    ) V(ColName, NewVal, OldVal);
  END

  -- UPDATE
  IF EXISTS(SELECT 1 FROM inserted) AND EXISTS(SELECT 1 FROM deleted)
  BEGIN
    INSERT dbo.AuditTrail (EventTimeUtc, SchemaName, TableName, PK, Operation, ColumnName, OldValue, NewValue,
                           SqlLogin, HostName, AppName, AppContext, Spid, RowHash, ChainHash)
    SELECT
      SYSUTCDATETIME(), @Schema, @Table,
      ' + @pkExpr_I + N' AS PK,
      ''U'' AS Operation,
      V.ColName,
      V.OldVal,
      V.NewVal,
      ORIGINAL_LOGIN(), HOST_NAME(), APP_NAME(), @ctx, @@SPID,
      HASHBYTES(''SHA1'', CONVERT(VARBINARY(MAX),
         ISNULL(V.ColName,'''')+ISNULL(V.OldVal,'''')+ISNULL(V.NewVal,'''')+ISNULL(@ctx,'''')
      )),
      HASHBYTES(''SHA1'',
         ISNULL((SELECT TOP(1) ChainHash FROM dbo.AuditTrail
                 WHERE TableName=@Table AND PK = (' + @pkExpr_I + N')
                 ORDER BY AuditId DESC), 0x) +
         CONVERT(VARBINARY(MAX),
           ISNULL(V.ColName,'''')+ISNULL(V.OldVal,'''')+ISNULL(V.NewVal,'''')+ISNULL(@ctx,'''')
         )
      )
    FROM inserted i
    JOIN  deleted d ON (' + @joinOn + N')
    CROSS APPLY (VALUES
' + @cols + N'
    ) V(ColName, NewVal, OldVal)
    WHERE ISNULL(V.NewVal, N''§NULL§'') <> ISNULL(V.OldVal, N''§NULL§'');
  END
END
';
  
  -- Clean up temp tables
  DROP TABLE #TempCols;
  DROP TABLE #TempPKs;
  
  EXEC (@trg);
END
