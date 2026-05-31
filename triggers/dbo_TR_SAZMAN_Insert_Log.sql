
CREATE TRIGGER [dbo].[TR_SAZMAN_Insert_Log]
ON [dbo].[SAZMAN]
AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @ApplicationName NVARCHAR(128);
    DECLARE @WindowsUsername NVARCHAR(128);
    DECLARE @ComputerName NVARCHAR(128);
    DECLARE @IPAddress VARCHAR(50);
    DECLARE @LoginName NVARCHAR(128);
    DECLARE @SessionID INT;
    DECLARE @SchemaName NVARCHAR(128);
    DECLARE @SPID INT;
    
    -- دریافت اطلاعات سشن
    SELECT 
        @ApplicationName = program_name,
        @WindowsUsername = SYSTEM_USER,
        @ComputerName = host_name,
        @LoginName = login_name,
        @SessionID = session_id,
        @SPID = @@SPID
    FROM sys.dm_exec_sessions
    WHERE session_id = @@SPID;
    
    SET @SchemaName = 'dbo';
    
    -- دریافت IP Address
    SELECT @IPAddress = client_net_address
    FROM sys.dm_exec_connections
    WHERE session_id = @@SPID;

    -- ثبت لاگ برای INSERT
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 
        'INSERT', 
        'FULL_ROW',
        NULL,
        'New Record Inserted',
        @ApplicationName, 
        @WindowsUsername, 
        @ComputerName, 
        @IPAddress, 
        @LoginName, 
        @SessionID, 
        @SchemaName, 
        @SPID, 
        @@TRANCOUNT
    FROM inserted i;
    
END
