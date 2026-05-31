
-- Create the trigger on the correct table
CREATE TRIGGER [dbo].[trg_Audit_YourMainTable]
ON [dbo].[STUF_FSK] -- The trigger is on STUF_FSK table
AFTER UPDATE, DELETE
AS
BEGIN
    SET NOCOUNT ON;

    -- === متغیرهایی برای نگهداری اطلاعات مشترک ===
    DECLARE @ActionType NVARCHAR(10);
    DECLARE @LoginName NVARCHAR(256);
    DECLARE @AppName NVARCHAR(128);
    DECLARE @HostName NVARCHAR(128);
    DECLARE @ClientIP VARCHAR(48);
    DECLARE @QueryText NVARCHAR(MAX);
    DECLARE @CurrentSPID INT = @@SPID;

    -- === دریافت متن کوئری اجرا شده ===
    CREATE TABLE #inputbuffer (
        EventType NVARCHAR(30),
        Parameters INT,
        EventInfo NVARCHAR(MAX)
    );
    INSERT INTO #inputbuffer
    EXEC ('DBCC INPUTBUFFER(' + @CurrentSPID + ')');
    SELECT @QueryText = EventInfo FROM #inputbuffer;
    DROP TABLE #inputbuffer;

    -- === دریافت اطلاعات نشست (Session) ===
    SELECT
        @LoginName = s.login_name,
        @AppName = s.program_name,
        @HostName = s.host_name,
        @ClientIP = c.client_net_address
    FROM sys.dm_exec_sessions AS s
    INNER JOIN sys.dm_exec_connections AS c ON s.session_id = c.session_id
    WHERE s.session_id = @CurrentSPID;


    -- === منطق اصلی تریگر: بررسی نوع عملیات ===

    -- 1. حالت DELETE
    IF NOT EXISTS (SELECT 1 FROM inserted)
    BEGIN
        SET @ActionType = 'DELETE';

        INSERT INTO dbo.AuditLogSTUF_FSK (
            SchemaName, TableName, RecordPK, ActionType, OldData, NewData,
            QueryText, EventTime, LoginName, AppName, HostName, ClientIP, SessionID
        )
        SELECT
            'dbo',
            'STUF_FSK',
            d.IDD,  -- <<<=== [اصلاح شد] نام کلید اصلی خود را اینجا قرار دهید
            @ActionType,
            (SELECT * FROM deleted WHERE IDD = d.IDD FOR XML RAW, TYPE, ELEMENTS), -- <<<=== [اصلاح شد]
            NULL,
            @QueryText,
            GETDATE(),
            @LoginName,
            @AppName,
            @HostName,
            @ClientIP,
            @CurrentSPID
        FROM deleted d;
    END

    -- 2. حالت UPDATE
    ELSE
    BEGIN
        SET @ActionType = 'UPDATE';

        INSERT INTO dbo.AuditLogSTUF_FSK (
            SchemaName, TableName, RecordPK, ActionType, OldData, NewData,
            QueryText, EventTime, LoginName, AppName, HostName, ClientIP, SessionID
        )
        SELECT
            'dbo',
            'STUF_FSK',
            i.IDD,  -- <<<=== [اصلاح شد] نام کلید اصلی خود را اینجا قرار دهید
            @ActionType,
            (SELECT * FROM deleted WHERE IDD = i.IDD FOR XML RAW, TYPE, ELEMENTS),  -- <<<=== [اصلاح شد]
            (SELECT * FROM inserted WHERE IDD = i.IDD FOR XML RAW, TYPE, ELEMENTS), -- <<<=== [اصلاح شد]
            @QueryText,
            GETDATE(),
            @LoginName,
            @AppName,
            @HostName,
            @ClientIP,
            @CurrentSPID
        FROM inserted i
        INNER JOIN deleted d ON i.IDD = d.IDD; -- <<<=== [اصلاح شد]
    END
END;
