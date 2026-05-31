
CREATE TRIGGER [dbo].[TR_SAZMAN_Update_Log]
ON [dbo].[SAZMAN]
AFTER UPDATE
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

    -- لاگ کردن تغییرات برای هر ستون
    
    -- UNIVERSITY_CO
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        d.UNIVERSITY_CO, 'UPDATE', 'UNIVERSITY_CO',
        CONVERT(NVARCHAR(MAX), d.UNIVERSITY_CO),
        CONVERT(NVARCHAR(MAX), i.UNIVERSITY_CO),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.UNIVERSITY_CO, 0) <> ISNULL(i.UNIVERSITY_CO, 0);

    -- NAME
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'NAME',
        d.NAME,
        i.NAME,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.NAME, '') <> ISNULL(i.NAME, '');

    -- CITY
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'CITY',
        d.CITY,
        i.CITY,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.CITY, '') <> ISNULL(i.CITY, '');

    -- MANAGER
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'MANAGER',
        d.MANAGER,
        i.MANAGER,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.MANAGER, '') <> ISNULL(i.MANAGER, '');

    -- MOAVEN
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'MOAVEN',
        d.MOAVEN,
        i.MOAVEN,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.MOAVEN, '') <> ISNULL(i.MOAVEN, '');

    -- ZIHESAB
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'ZIHESAB',
        d.ZIHESAB,
        i.ZIHESAB,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.ZIHESAB, '') <> ISNULL(i.ZIHESAB, '');

    -- AMINAMVAL
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'AMINAMVAL',
        d.AMINAMVAL,
        i.AMINAMVAL,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.AMINAMVAL, '') <> ISNULL(i.AMINAMVAL, '');

    -- YEA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'YEA',
        CONVERT(NVARCHAR(MAX), d.YEA),
        CONVERT(NVARCHAR(MAX), i.YEA),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.YEA, 0) <> ISNULL(i.YEA, 0);

    -- SANAD
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SANAD',
        CONVERT(NVARCHAR(MAX), d.SANAD),
        CONVERT(NVARCHAR(MAX), i.SANAD),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SANAD, 0) <> ISNULL(i.SANAD, 0);

    -- GHAYM
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'GHAYM',
        CONVERT(NVARCHAR(MAX), d.GHAYM),
        CONVERT(NVARCHAR(MAX), i.GHAYM),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.GHAYM, 0) <> ISNULL(i.GHAYM, 0);

    -- KALA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'KALA',
        CONVERT(NVARCHAR(MAX), d.KALA),
        CONVERT(NVARCHAR(MAX), i.KALA),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.KALA, 0) <> ISNULL(i.KALA, 0);

    -- PERSON
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PERSON',
        CONVERT(NVARCHAR(MAX), d.PERSON),
        CONVERT(NVARCHAR(MAX), i.PERSON),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PERSON, 0) <> ISNULL(i.PERSON, 0);

    -- DIG
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'DIG',
        CONVERT(NVARCHAR(MAX), d.DIG),
        CONVERT(NVARCHAR(MAX), i.DIG),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.DIG, 0) <> ISNULL(i.DIG, 0);

    -- WAR
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'WAR',
        CONVERT(NVARCHAR(MAX), d.WAR),
        CONVERT(NVARCHAR(MAX), i.WAR),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.WAR, 0) <> ISNULL(i.WAR, 0);

    -- LST
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'LST',
        CONVERT(NVARCHAR(MAX), d.LST),
        CONVERT(NVARCHAR(MAX), i.LST),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.LST, 0) <> ISNULL(i.LST, 0);

    -- TFTPAGE
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'TFTPAGE',
        CONVERT(NVARCHAR(MAX), d.TFTPAGE),
        CONVERT(NVARCHAR(MAX), i.TFTPAGE),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.TFTPAGE, 0) <> ISNULL(i.TFTPAGE, 0);

    -- TFSAZMAN
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'TFSAZMAN',
        d.TFSAZMAN,
        i.TFSAZMAN,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.TFSAZMAN, '') <> ISNULL(i.TFSAZMAN, '');

    -- TFADDRESS
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'TFADDRESS',
        d.TFADDRESS,
        i.TFADDRESS,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.TFADDRESS, '') <> ISNULL(i.TFADDRESS, '');

    -- TFTEL
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'TFTEL',
        d.TFTEL,
        i.TFTEL,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.TFTEL, '') <> ISNULL(i.TFTEL, '');

    -- TFCODE_E
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'TFCODE_E',
        d.TFCODE_E,
        i.TFCODE_E,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.TFCODE_E, '') <> ISNULL(i.TFCODE_E, '');

    -- WIDTH_D
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'WIDTH_D',
        d.WIDTH_D,
        i.WIDTH_D,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.WIDTH_D, '') <> ISNULL(i.WIDTH_D, '');

    -- HIGH_D
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HIGH_D',
        d.HIGH_D,
        i.HIGH_D,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HIGH_D, '') <> ISNULL(i.HIGH_D, '');

    -- CPI
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'CPI',
        CONVERT(NVARCHAR(MAX), d.CPI),
        CONVERT(NVARCHAR(MAX), i.CPI),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.CPI, 0) <> ISNULL(i.CPI, 0);

    -- SANDOGH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SANDOGH',
        CONVERT(NVARCHAR(MAX), d.SANDOGH),
        CONVERT(NVARCHAR(MAX), i.SANDOGH),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SANDOGH, 0) <> ISNULL(i.SANDOGH, 0);

    -- BANKHA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'BANKHA',
        CONVERT(NVARCHAR(MAX), d.BANKHA),
        CONVERT(NVARCHAR(MAX), i.BANKHA),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.BANKHA, 0) <> ISNULL(i.BANKHA, 0);

    -- BESTANKAR
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'BESTANKAR',
        CONVERT(NVARCHAR(MAX), d.BESTANKAR),
        CONVERT(NVARCHAR(MAX), i.BESTANKAR),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.BESTANKAR, 0) <> ISNULL(i.BESTANKAR, 0);

    -- BEDEHKAR
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'BEDEHKAR',
        CONVERT(NVARCHAR(MAX), d.BEDEHKAR),
        CONVERT(NVARCHAR(MAX), i.BEDEHKAR),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.BEDEHKAR, 0) <> ISNULL(i.BEDEHKAR, 0);

    -- KHARID
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'KHARID',
        CONVERT(NVARCHAR(MAX), d.KHARID),
        CONVERT(NVARCHAR(MAX), i.KHARID),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.KHARID, 0) <> ISNULL(i.KHARID, 0);

    -- MKHARID
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'MKHARID',
        CONVERT(NVARCHAR(MAX), d.MKHARID),
        CONVERT(NVARCHAR(MAX), i.MKHARID),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.MKHARID, 0) <> ISNULL(i.MKHARID, 0);

    -- TKHARID
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'TKHARID',
        CONVERT(NVARCHAR(MAX), d.TKHARID),
        CONVERT(NVARCHAR(MAX), i.TKHARID),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.TKHARID, 0) <> ISNULL(i.TKHARID, 0);

    -- HKHARID
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HKHARID',
        CONVERT(NVARCHAR(MAX), d.HKHARID),
        CONVERT(NVARCHAR(MAX), i.HKHARID),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HKHARID, 0) <> ISNULL(i.HKHARID, 0);

    -- FROSH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'FROSH',
        CONVERT(NVARCHAR(MAX), d.FROSH),
        CONVERT(NVARCHAR(MAX), i.FROSH),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.FROSH, 0) <> ISNULL(i.FROSH, 0);

    -- MFROSH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'MFROSH',
        CONVERT(NVARCHAR(MAX), d.MFROSH),
        CONVERT(NVARCHAR(MAX), i.MFROSH),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.MFROSH, 0) <> ISNULL(i.MFROSH, 0);

    -- TFROSH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'TFROSH',
        CONVERT(NVARCHAR(MAX), d.TFROSH),
        CONVERT(NVARCHAR(MAX), i.TFROSH),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.TFROSH, 0) <> ISNULL(i.TFROSH, 0);

    -- HFROSH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HFROSH',
        CONVERT(NVARCHAR(MAX), d.HFROSH),
        CONVERT(NVARCHAR(MAX), i.HFROSH),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HFROSH, 0) <> ISNULL(i.HFROSH, 0);

    -- MOGODIA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'MOGODIA',
        CONVERT(NVARCHAR(MAX), d.MOGODIA),
        CONVERT(NVARCHAR(MAX), i.MOGODIA),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.MOGODIA, 0) <> ISNULL(i.MOGODIA, 0);

    -- MOGODIP
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'MOGODIP',
        CONVERT(NVARCHAR(MAX), d.MOGODIP),
        CONVERT(NVARCHAR(MAX), i.MOGODIP),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.MOGODIP, 0) <> ISNULL(i.MOGODIP, 0);

    -- DARAM
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'DARAM',
        CONVERT(NVARCHAR(MAX), d.DARAM),
        CONVERT(NVARCHAR(MAX), i.DARAM),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.DARAM, 0) <> ISNULL(i.DARAM, 0);

    -- HDARAM
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HDARAM',
        CONVERT(NVARCHAR(MAX), d.HDARAM),
        CONVERT(NVARCHAR(MAX), i.HDARAM),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HDARAM, 0) <> ISNULL(i.HDARAM, 0);

    -- HKOL
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HKOL',
        CONVERT(NVARCHAR(MAX), d.HKOL),
        CONVERT(NVARCHAR(MAX), i.HKOL),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HKOL, 0) <> ISNULL(i.HKOL, 0);

    -- ADA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'ADA',
        d.ADA,
        i.ADA,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.ADA, '') <> ISNULL(i.ADA, '');

    -- APA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'APA',
        d.APA,
        i.APA,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.APA, '') <> ISNULL(i.APA, '');

    -- ADV
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'ADV',
        d.ADV,
        i.ADV,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.ADV, '') <> ISNULL(i.ADV, '');

    -- HAVALAH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HAVALAH',
        CONVERT(NVARCHAR(MAX), d.HAVALAH),
        CONVERT(NVARCHAR(MAX), i.HAVALAH),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HAVALAH, 0) <> ISNULL(i.HAVALAH, 0);

    -- CTRL_TS
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'CTRL_TS',
        CONVERT(NVARCHAR(MAX), d.CTRL_TS),
        CONVERT(NVARCHAR(MAX), i.CTRL_TS),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.CTRL_TS, 0) <> ISNULL(i.CTRL_TS, 0);

    -- F_ANBARF
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'F_ANBARF',
        CONVERT(NVARCHAR(MAX), d.F_ANBARF),
        CONVERT(NVARCHAR(MAX), i.F_ANBARF),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.F_ANBARF, 0) <> ISNULL(i.F_ANBARF, 0);

    -- GH_PK
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'GH_PK',
        CONVERT(NVARCHAR(MAX), d.GH_PK),
        CONVERT(NVARCHAR(MAX), i.GH_PK),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.GH_PK, 0) <> ISNULL(i.GH_PK, 0);

    -- L_NUMBER
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'L_NUMBER',
        CONVERT(NVARCHAR(MAX), d.L_NUMBER),
        CONVERT(NVARCHAR(MAX), i.L_NUMBER),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.L_NUMBER, 0) <> ISNULL(i.L_NUMBER, 0);

    -- SF_G
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SF_G',
        CONVERT(NVARCHAR(MAX), d.SF_G),
        CONVERT(NVARCHAR(MAX), i.SF_G),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SF_G, 0) <> ISNULL(i.SF_G, 0);

    -- TAR_KM
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'TAR_KM',
        CONVERT(NVARCHAR(MAX), d.TAR_KM),
        CONVERT(NVARCHAR(MAX), i.TAR_KM),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.TAR_KM, 0) <> ISNULL(i.TAR_KM, 0);

    -- BACKPATH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'BACKPATH',
        d.BACKPATH,
        i.BACKPATH,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.BACKPATH, '') <> ISNULL(i.BACKPATH, '');

    -- TKHF
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'TKHF',
        CONVERT(NVARCHAR(MAX), d.TKHF),
        CONVERT(NVARCHAR(MAX), i.TKHF),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.TKHF, 0) <> ISNULL(i.TKHF, 0);

    -- HAZ_TOL
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HAZ_TOL',
        CONVERT(NVARCHAR(MAX), d.HAZ_TOL),
        CONVERT(NVARCHAR(MAX), i.HAZ_TOL),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HAZ_TOL, 0) <> ISNULL(i.HAZ_TOL, 0);

    -- PJHAZ_TOL1
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PJHAZ_TOL1',
        CONVERT(NVARCHAR(MAX), d.PJHAZ_TOL1),
        CONVERT(NVARCHAR(MAX), i.PJHAZ_TOL1),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PJHAZ_TOL1, 0) <> ISNULL(i.PJHAZ_TOL1, 0);

    -- PHAZ_TOL
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PHAZ_TOL',
        CONVERT(NVARCHAR(MAX), d.PHAZ_TOL),
        CONVERT(NVARCHAR(MAX), i.PHAZ_TOL),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PHAZ_TOL, 0) <> ISNULL(i.PHAZ_TOL, 0);

    -- GHEYMAT
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'GHEYMAT',
        CONVERT(NVARCHAR(MAX), d.GHEYMAT),
        CONVERT(NVARCHAR(MAX), i.GHEYMAT),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.GHEYMAT, 0) <> ISNULL(i.GHEYMAT, 0);

    -- PPDAST
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PPDAST',
        CONVERT(NVARCHAR(MAX), d.PPDAST),
        CONVERT(NVARCHAR(MAX), i.PPDAST),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PPDAST, 0) <> ISNULL(i.PPDAST, 0);

    -- PPSAR
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PPSAR',
        CONVERT(NVARCHAR(MAX), d.PPSAR),
        CONVERT(NVARCHAR(MAX), i.PPSAR),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PPSAR, 0) <> ISNULL(i.PPSAR, 0);

    -- AMALKARD
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'AMALKARD',
        CONVERT(NVARCHAR(MAX), d.AMALKARD),
        CONVERT(NVARCHAR(MAX), i.AMALKARD),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.AMALKARD, 0) <> ISNULL(i.AMALKARD, 0);

    -- PERSONEL
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PERSONEL',
        d.PERSONEL,
        i.PERSONEL,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PERSONEL, '') <> ISNULL(i.PERSONEL, '');

    -- PERVAM
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PERVAM',
        d.PERVAM,
        i.PERVAM,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PERVAM, '') <> ISNULL(i.PERVAM, '');

    -- CONKAL
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'CONKAL',
        CONVERT(NVARCHAR(MAX), d.CONKAL),
        CONVERT(NVARCHAR(MAX), i.CONKAL),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.CONKAL, 0) <> ISNULL(i.CONKAL, 0);

    -- HNAH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HNAH',
        CONVERT(NVARCHAR(MAX), CAST(d.HNAH AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.HNAH AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.HNAH <> i.HNAH;

    -- HEZA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HEZA',
        CONVERT(NVARCHAR(MAX), CAST(d.HEZA AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.HEZA AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.HEZA <> i.HEZA;

    -- HPAD
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HPAD',
        CONVERT(NVARCHAR(MAX), CAST(d.HPAD AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.HPAD AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.HPAD <> i.HPAD;

    -- HOLA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HOLA',
        CONVERT(NVARCHAR(MAX), CAST(d.HOLA AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.HOLA AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.HOLA <> i.HOLA;

    -- HKHA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HKHA',
        CONVERT(NVARCHAR(MAX), CAST(d.HKHA AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.HKHA AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.HKHA <> i.HKHA;

    -- HJAZ
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HJAZ',
        CONVERT(NVARCHAR(MAX), CAST(d.HJAZ AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.HJAZ AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.HJAZ <> i.HJAZ;

    -- HRAN
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HRAN',
        CONVERT(NVARCHAR(MAX), CAST(d.HRAN AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.HRAN AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.HRAN <> i.HRAN;

    -- HSAY
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HSAY',
        CONVERT(NVARCHAR(MAX), CAST(d.HSAY AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.HSAY AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.HSAY <> i.HSAY;

    -- HCON
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HCON',
        CONVERT(NVARCHAR(MAX), CAST(d.HCON AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.HCON AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.HCON <> i.HCON;

    -- HSHI
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HSHI',
        CONVERT(NVARCHAR(MAX), CAST(d.HSHI AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.HSHI AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.HSHI <> i.HSHI;

    -- ادامه برای بقیه ستون‌ها به همین شکل...
    -- (به دلیل محدودیت طول، بقیه ستون‌ها را به همین روال ادامه دهید)

    -- HAZEDAR
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HAZEDAR',
        d.HAZEDAR,
        i.HAZEDAR,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HAZEDAR, '') <> ISNULL(i.HAZEDAR, '');

    -- EDABIM
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'EDABIM',
        d.EDABIM,
        i.EDABIM,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.EDABIM, '') <> ISNULL(i.EDABIM, '');

    -- برای صرفه‌جویی در فضا، الگو را برای بقیه ستون‌ها تکرار کنید...
    -- این تریگر باید برای تمام 200+ ستون ادامه یابد

-- ادامه تریگر UPDATE از جایی که قطع شده بود

    -- HAZBIM
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HAZBIM',
        d.HAZBIM,
        i.HAZBIM,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HAZBIM, '') <> ISNULL(i.HAZBIM, '');

    -- BESHO
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'BESHO',
        d.BESHO,
        i.BESHO,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.BESHO, '') <> ISNULL(i.BESHO, '');

    -- BEDMOS
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'BEDMOS',
        d.BEDMOS,
        i.BEDMOS,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.BEDMOS, '') <> ISNULL(i.BEDMOS, '');

    -- PARDAKH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PARDAKH',
        d.PARDAKH,
        i.PARDAKH,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PARDAKH, '') <> ISNULL(i.PARDAKH, '');

    -- HAZMALI
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HAZMALI',
        d.HAZMALI,
        i.HAZMALI,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HAZMALI, '') <> ISNULL(i.HAZMALI, '');

    -- SAGHFH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SAGHFH',
        CONVERT(NVARCHAR(MAX), d.SAGHFH),
        CONVERT(NVARCHAR(MAX), i.SAGHFH),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SAGHFH, 0) <> ISNULL(i.SAGHFH, 0);

    -- MAND
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'MAND',
        CONVERT(NVARCHAR(MAX), CAST(d.MAND AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.MAND AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.MAND <> i.MAND;

    -- MOJU
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'MOJU',
        CONVERT(NVARCHAR(MAX), CAST(d.MOJU AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.MOJU AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.MOJU <> i.MOJU;

    -- SA_HOGH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SA_HOGH',
        CONVERT(NVARCHAR(MAX), CAST(d.SA_HOGH AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SA_HOGH AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.SA_HOGH <> i.SA_HOGH;

    -- SA_40EZ
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SA_40EZ',
        CONVERT(NVARCHAR(MAX), CAST(d.SA_40EZ AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SA_40EZ AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.SA_40EZ <> i.SA_40EZ;

    -- SA_EZAF
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SA_EZAF',
        CONVERT(NVARCHAR(MAX), CAST(d.SA_EZAF AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SA_EZAF AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.SA_EZAF <> i.SA_EZAF;

    -- SA_PADA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SA_PADA',
        CONVERT(NVARCHAR(MAX), CAST(d.SA_PADA AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SA_PADA AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.SA_PADA <> i.SA_PADA;

    -- SA_HOLA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SA_HOLA',
        CONVERT(NVARCHAR(MAX), CAST(d.SA_HOLA AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SA_HOLA AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.SA_HOLA <> i.SA_HOLA;

    -- SA_KHAR
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SA_KHAR',
        CONVERT(NVARCHAR(MAX), CAST(d.SA_KHAR AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SA_KHAR AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.SA_KHAR <> i.SA_KHAR;

    -- SA_NAHA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SA_NAHA',
        CONVERT(NVARCHAR(MAX), CAST(d.SA_NAHA AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SA_NAHA AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.SA_NAHA <> i.SA_NAHA;

    -- SA_JAZB
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SA_JAZB',
        CONVERT(NVARCHAR(MAX), CAST(d.SA_JAZB AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SA_JAZB AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.SA_JAZB <> i.SA_JAZB;

    -- SA_RAND
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SA_RAND',
        CONVERT(NVARCHAR(MAX), CAST(d.SA_RAND AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SA_RAND AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.SA_RAND <> i.SA_RAND;

    -- SA_COND
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SA_COND',
        CONVERT(NVARCHAR(MAX), CAST(d.SA_COND AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SA_COND AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.SA_COND <> i.SA_COND;

    -- SA_SAYE
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SA_SAYE',
        CONVERT(NVARCHAR(MAX), CAST(d.SA_SAYE AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SA_SAYE AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.SA_SAYE <> i.SA_SAYE;

    -- SA_23BI
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SA_23BI',
        CONVERT(NVARCHAR(MAX), CAST(d.SA_23BI AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SA_23BI AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.SA_23BI <> i.SA_23BI;

    -- HAZTOLID
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HAZTOLID',
        d.HAZTOLID,
        i.HAZTOLID,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HAZTOLID, '') <> ISNULL(i.HAZTOLID, '');

    -- HAZFROOSH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HAZFROOSH',
        d.HAZFROOSH,
        i.HAZFROOSH,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HAZFROOSH, '') <> ISNULL(i.HAZFROOSH, '');

    -- HAZKHADAMAT
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HAZKHADAMAT',
        d.HAZKHADAMAT,
        i.HAZKHADAMAT,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HAZKHADAMAT, '') <> ISNULL(i.HAZKHADAMAT, '');

    -- PISHDAR
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PISHDAR',
        CONVERT(NVARCHAR(MAX), d.PISHDAR),
        CONVERT(NVARCHAR(MAX), i.PISHDAR),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PISHDAR, 0) <> ISNULL(i.PISHDAR, 0);

    -- DEFANB
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'DEFANB',
        CONVERT(NVARCHAR(MAX), d.DEFANB),
        CONVERT(NVARCHAR(MAX), i.DEFANB),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.DEFANB, 0) <> ISNULL(i.DEFANB, 0);

    -- DEFTKH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'DEFTKH',
        CONVERT(NVARCHAR(MAX), d.DEFTKH),
        CONVERT(NVARCHAR(MAX), i.DEFTKH),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.DEFTKH, 0) <> ISNULL(i.DEFTKH, 0);

    -- ECONM
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'ECONM',
        CONVERT(NVARCHAR(MAX), CAST(d.ECONM AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.ECONM AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.ECONM, 0) <> ISNULL(i.ECONM, 0);

    -- FRUP
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'FRUP',
        CONVERT(NVARCHAR(MAX), CAST(d.FRUP AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.FRUP AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.FRUP, 0) <> ISNULL(i.FRUP, 0);

    -- UPDDATE
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'UPDDATE',
        CONVERT(NVARCHAR(MAX), CAST(d.UPDDATE AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.UPDDATE AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.UPDDATE, 0) <> ISNULL(i.UPDDATE, 0);

    -- FINALS
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'FINALS',
        CONVERT(NVARCHAR(MAX), CAST(d.FINALS AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.FINALS AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.FINALS, 0) <> ISNULL(i.FINALS, 0);

    -- PSANDHES
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PSANDHES',
        d.PSANDHES,
        i.PSANDHES,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PSANDHES, '') <> ISNULL(i.PSANDHES, '');

    -- SANAVP
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SANAVP',
        CONVERT(NVARCHAR(MAX), d.SANAVP),
        CONVERT(NVARCHAR(MAX), i.SANAVP),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SANAVP, 0) <> ISNULL(i.SANAVP, 0);

    -- BON
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'BON',
        CONVERT(NVARCHAR(MAX), d.BON),
        CONVERT(NVARCHAR(MAX), i.BON),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.BON, 0) <> ISNULL(i.BON, 0);

    -- ISO_FROOSH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'ISO_FROOSH',
        d.ISO_FROOSH,
        i.ISO_FROOSH,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.ISO_FROOSH, '') <> ISNULL(i.ISO_FROOSH, '');

    -- ISO_KHAREED
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'ISO_KHAREED',
        d.ISO_KHAREED,
        i.ISO_KHAREED,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.ISO_KHAREED, '') <> ISNULL(i.ISO_KHAREED, '');

    -- ISO_MAVAD
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'ISO_MAVAD',
        d.ISO_MAVAD,
        i.ISO_MAVAD,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.ISO_MAVAD, '') <> ISNULL(i.ISO_MAVAD, '');

    -- ISO_TOLID
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'ISO_TOLID',
        d.ISO_TOLID,
        i.ISO_TOLID,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.ISO_TOLID, '') <> ISNULL(i.ISO_TOLID, '');

    -- ISO_MAVADSAYER
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'ISO_MAVADSAYER',
        d.ISO_MAVADSAYER,
        i.ISO_MAVADSAYER,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.ISO_MAVADSAYER, '') <> ISNULL(i.ISO_MAVADSAYER, '');

    -- SANAT
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SANAT',
        CONVERT(NVARCHAR(MAX), CAST(d.SANAT AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SANAT AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SANAT, 0) <> ISNULL(i.SANAT, 0);

    -- CODEVIEW
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'CODEVIEW',
        CONVERT(NVARCHAR(MAX), d.CODEVIEW),
        CONVERT(NVARCHAR(MAX), i.CODEVIEW),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.CODEVIEW, 0) <> ISNULL(i.CODEVIEW, 0);

    -- PKHARID
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PKHARID',
        CONVERT(NVARCHAR(MAX), d.PKHARID),
        CONVERT(NVARCHAR(MAX), i.PKHARID),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PKHARID, 0) <> ISNULL(i.PKHARID, 0);

    -- SIGN
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SIGN',
        CONVERT(NVARCHAR(MAX), CAST(d.SIGN AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SIGN AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SIGN, 0) <> ISNULL(i.SIGN, 0);

    -- BARCOD
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'BARCOD',
        CONVERT(NVARCHAR(MAX), CAST(d.BARCOD AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.BARCOD AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.BARCOD, 0) <> ISNULL(i.BARCOD, 0);

    -- SAGHF
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SAGHF',
        CONVERT(NVARCHAR(MAX), CAST(d.SAGHF AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SAGHF AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SAGHF, 0) <> ISNULL(i.SAGHF, 0);

    -- SERVERNAM
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SERVERNAM',
        d.SERVERNAM,
        i.SERVERNAM,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SERVERNAM, '') <> ISNULL(i.SERVERNAM, '');

    -- TENDAR
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'TENDAR',
        CONVERT(NVARCHAR(MAX), CAST(d.TENDAR AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.TENDAR AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.TENDAR, 0) <> ISNULL(i.TENDAR, 0);

    -- LECOL1
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'LECOL1',
        d.LECOL1,
        i.LECOL1,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.LECOL1, '') <> ISNULL(i.LECOL1, '');

    -- LECOL2
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'LECOL2',
        d.LECOL2,
        i.LECOL2,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.LECOL2, '') <> ISNULL(i.LECOL2, '');

    -- LECOL3
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'LECOL3',
        d.LECOL3,
        i.LECOL3,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.LECOL3, '') <> ISNULL(i.LECOL3, '');

    -- LECOL4
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'LECOL4',
        d.LECOL4,
        i.LECOL4,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.LECOL4, '') <> ISNULL(i.LECOL4, '');

    -- LKCOL1
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'LKCOL1',
        d.LKCOL1,
        i.LKCOL1,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.LKCOL1, '') <> ISNULL(i.LKCOL1, '');

    -- HESMBAA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HESMBAA',
        d.HESMBAA,
        i.HESMBAA,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HESMBAA, '') <> ISNULL(i.HESMBAA, '');

    -- ECODE
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'ECODE',
        d.ECODE,
        i.ECODE,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.ECODE, '') <> ISNULL(i.ECODE, '');

    -- PCODE
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PCODE',
        d.PCODE,
        i.PCODE,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PCODE, '') <> ISNULL(i.PCODE, '');

    -- IYALAT
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IYALAT',
        d.IYALAT,
        i.IYALAT,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IYALAT, '') <> ISNULL(i.IYALAT, '');

    -- MCODEM
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'MCODEM',
        d.MCODEM,
        i.MCODEM,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.MCODEM, '') <> ISNULL(i.MCODEM, '');

    -- HPOR
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HPOR',
        d.HPOR,
        i.HPOR,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HPOR, '') <> ISNULL(i.HPOR, '');

    -- SAGHF2
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SAGHF2',
        CONVERT(NVARCHAR(MAX), CAST(d.SAGHF2 AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SAGHF2 AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SAGHF2, 0) <> ISNULL(i.SAGHF2, 0);

    -- OPTIONSS
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'OPTIONSS',
        d.OPTIONSS,
        i.OPTIONSS,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.OPTIONSS, '') <> ISNULL(i.OPTIONSS, '');

    -- CTL_DT
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'CTL_DT',
        CONVERT(NVARCHAR(MAX), CAST(d.CTL_DT AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.CTL_DT AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.CTL_DT, 0) <> ISNULL(i.CTL_DT, 0);

    -- LOCKFAP
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'LOCKFAP',
        CONVERT(NVARCHAR(MAX), CAST(d.LOCKFAP AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.LOCKFAP AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.LOCKFAP, 0) <> ISNULL(i.LOCKFAP, 0);

    -- LOCKFSI
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'LOCKFSI',
        CONVERT(NVARCHAR(MAX), CAST(d.LOCKFSI AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.LOCKFSI AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.LOCKFSI, 0) <> ISNULL(i.LOCKFSI, 0);

    -- TRANSF
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'TRANSF',
        CONVERT(NVARCHAR(MAX), CAST(d.TRANSF AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.TRANSF AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.TRANSF, 0) <> ISNULL(i.TRANSF, 0);

    -- OKF
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'OKF',
        CONVERT(NVARCHAR(MAX), CAST(d.OKF AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.OKF AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.OKF, 0) <> ISNULL(i.OKF, 0);

    -- ARSESH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'ARSESH',
        CONVERT(NVARCHAR(MAX), d.ARSESH),
        CONVERT(NVARCHAR(MAX), i.ARSESH),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.ARSESH, 0) <> ISNULL(i.ARSESH, 0);

    -- RMOG
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'RMOG',
        CONVERT(NVARCHAR(MAX), CAST(d.RMOG AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.RMOG AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.RMOG, 0) <> ISNULL(i.RMOG, 0);

    -- APV
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'APV',
        d.APV,
        i.APV,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.APV, '') <> ISNULL(i.APV, '');

    -- HOTCOD
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HOTCOD',
        d.HOTCOD,
        i.HOTCOD,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HOTCOD, '') <> ISNULL(i.HOTCOD, '');

    -- STFR
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'STFR',
        CONVERT(NVARCHAR(MAX), d.STFR),
        CONVERT(NVARCHAR(MAX), i.STFR),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.STFR, 0) <> ISNULL(i.STFR, 0);

    -- STKH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'STKH',
        CONVERT(NVARCHAR(MAX), d.STKH),
        CONVERT(NVARCHAR(MAX), i.STKH),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.STKH, 0) <> ISNULL(i.STKH, 0);

    -- STHFR
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'STHFR',
        CONVERT(NVARCHAR(MAX), d.STHFR),
        CONVERT(NVARCHAR(MAX), i.STHFR),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.STHFR, 0) <> ISNULL(i.STHFR, 0);

    -- STHKH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'STHKH',
        CONVERT(NVARCHAR(MAX), d.STHKH),
        CONVERT(NVARCHAR(MAX), i.STHKH),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.STHKH, 0) <> ISNULL(i.STHKH, 0);

    -- STENT
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'STENT',
        CONVERT(NVARCHAR(MAX), d.STENT),
        CONVERT(NVARCHAR(MAX), i.STENT),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.STENT, 0) <> ISNULL(i.STENT, 0);

    -- STKHS
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'STKHS',
        CONVERT(NVARCHAR(MAX), d.STKHS),
        CONVERT(NVARCHAR(MAX), i.STKHS),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.STKHS, 0) <> ISNULL(i.STKHS, 0);

    -- STKHH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'STKHH',
        CONVERT(NVARCHAR(MAX), d.STKHH),
        CONVERT(NVARCHAR(MAX), i.STKHH),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.STKHH, 0) <> ISNULL(i.STKHH, 0);

    -- STTOL
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'STTOL',
        CONVERT(NVARCHAR(MAX), d.STTOL),
        CONVERT(NVARCHAR(MAX), i.STTOL),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.STTOL, 0) <> ISNULL(i.STTOL, 0);

    -- STFRB
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'STFRB',
        CONVERT(NVARCHAR(MAX), d.STFRB),
        CONVERT(NVARCHAR(MAX), i.STFRB),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.STFRB, 0) <> ISNULL(i.STFRB, 0);

    -- STBKH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'STBKH',
        CONVERT(NVARCHAR(MAX), d.STBKH),
        CONVERT(NVARCHAR(MAX), i.STBKH),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.STBKH, 0) <> ISNULL(i.STBKH, 0);

    -- STMO
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'STMO',
        CONVERT(NVARCHAR(MAX), d.STMO),
        CONVERT(NVARCHAR(MAX), i.STMO),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.STMO, 0) <> ISNULL(i.STMO, 0);

    -- STKHA
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'STKHA',
        CONVERT(NVARCHAR(MAX), d.STKHA),
        CONVERT(NVARCHAR(MAX), i.STKHA),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.STKHA, 0) <> ISNULL(i.STKHA, 0);

    -- SNDKH
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SNDKH',
        CONVERT(NVARCHAR(MAX), CAST(d.SNDKH AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SNDKH AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SNDKH, 0) <> ISNULL(i.SNDKH, 0);

    -- SMS_USERNAME
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_USERNAME',
        d.SMS_USERNAME,
        i.SMS_USERNAME,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_USERNAME, '') <> ISNULL(i.SMS_USERNAME, '');

    -- SMS_PASSWORD
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_PASSWORD',
        d.SMS_PASSWORD,
        i.SMS_PASSWORD,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_PASSWORD, '') <> ISNULL(i.SMS_PASSWORD, '');

    -- SMS_LIBKEY
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_LIBKEY',
        d.SMS_LIBKEY,
        i.SMS_LIBKEY,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_LIBKEY, '') <> ISNULL(i.SMS_LIBKEY, '');

    -- SMS_TSMSHOST
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_TSMSHOST',
        d.SMS_TSMSHOST,
        i.SMS_TSMSHOST,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_TSMSHOST, '') <> ISNULL(i.SMS_TSMSHOST, '');

    -- SMS_ProxyUserName
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_ProxyUserName',
        d.SMS_ProxyUserName,
        i.SMS_ProxyUserName,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_ProxyUserName, '') <> ISNULL(i.SMS_ProxyUserName, '');

    -- SMS_ProxyPassword
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_ProxyPassword',
        d.SMS_ProxyPassword,
        i.SMS_ProxyPassword,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_ProxyPassword, '') <> ISNULL(i.SMS_ProxyPassword, '');

    -- SMS_ProxyServer
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_ProxyServer',
        d.SMS_ProxyServer,
        i.SMS_ProxyServer,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_ProxyServer, '') <> ISNULL(i.SMS_ProxyServer, '');

    -- SMS_ProxyPort
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_ProxyPort',
        d.SMS_ProxyPort,
        i.SMS_ProxyPort,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_ProxyPort, '') <> ISNULL(i.SMS_ProxyPort, '');

    -- SMS_FirewallUserName
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_FirewallUserName',
        d.SMS_FirewallUserName,
        i.SMS_FirewallUserName,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_FirewallUserName, '') <> ISNULL(i.SMS_FirewallUserName, '');

    -- SMS_FirewallPassword
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_FirewallPassword',
        d.SMS_FirewallPassword,
        i.SMS_FirewallPassword,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_FirewallPassword, '') <> ISNULL(i.SMS_FirewallPassword, '');

    -- SMS_FirewallHost
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_FirewallHost',
        d.SMS_FirewallHost,
        i.SMS_FirewallHost,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_FirewallHost, '') <> ISNULL(i.SMS_FirewallHost, '');

    -- SMS_FirewallPort
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_FirewallPort',
        d.SMS_FirewallPort,
        i.SMS_FirewallPort,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_FirewallPort, '') <> ISNULL(i.SMS_FirewallPort, '');

    -- SMS_FirewallType
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_FirewallType',
        d.SMS_FirewallType,
        i.SMS_FirewallType,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_FirewallType, '') <> ISNULL(i.SMS_FirewallType, '');

    -- DSMS
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'DSMS',
        CONVERT(NVARCHAR(MAX), CAST(d.DSMS AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.DSMS AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.DSMS, 0) <> ISNULL(i.DSMS, 0);

    -- SMS_OWNER
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMS_OWNER',
        d.SMS_OWNER,
        i.SMS_OWNER,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMS_OWNER, '') <> ISNULL(i.SMS_OWNER, '');

    -- PRMFR
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PRMFR',
        CONVERT(NVARCHAR(MAX), CAST(d.PRMFR AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.PRMFR AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PRMFR, 0) <> ISNULL(i.PRMFR, 0);

    -- SMSACT
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMSACT',
        CONVERT(NVARCHAR(MAX), CAST(d.SMSACT AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SMSACT AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMSACT, 0) <> ISNULL(i.SMSACT, 0);

    -- HESDESK
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HESDESK',
        d.HESDESK,
        i.HESDESK,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HESDESK, '') <> ISNULL(i.HESDESK, '');

    -- ISO_DTOLID
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'ISO_DTOLID',
        d.ISO_DTOLID,
        i.ISO_DTOLID,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.ISO_DTOLID, '') <> ISNULL(i.ISO_DTOLID, '');

    -- SERFACB
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SERFACB',
        CONVERT(NVARCHAR(MAX), CAST(d.SERFACB AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.SERFACB AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SERFACB, 0) <> ISNULL(i.SERFACB, 0);

    -- HBON
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HBON',
        CONVERT(NVARCHAR(MAX), CAST(d.HBON AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.HBON AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HBON, 0) <> ISNULL(i.HBON, 0);

    -- pishpross
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'pishpross',
        CONVERT(NVARCHAR(MAX), d.pishpross),
        CONVERT(NVARCHAR(MAX), i.pishpross),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.pishpross, 0) <> ISNULL(i.pishpross, 0);

    -- version
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'version',
        d.version,
        i.version,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.version, '') <> ISNULL(i.version, '');

    -- hesnaghd
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'hesnaghd',
        d.hesnaghd,
        i.hesnaghd,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, 
        @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.hesnaghd, '') <> ISNULL(i.hesnaghd, '');

-- ادامه تریگر UPDATE - فیلدهای IT و IS و سایر موارد باقیمانده

    -- IT1
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IT1',
        d.IT1,
        i.IT1,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IT1, '') <> ISNULL(i.IT1, '');

    -- IT2
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IT2',
        d.IT2,
        i.IT2,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IT2, '') <> ISNULL(i.IT2, '');

    -- IT3
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IT3',
        d.IT3,
        i.IT3,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IT3, '') <> ISNULL(i.IT3, '');

    -- IT4
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IT4',
        d.IT4,
        i.IT4,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IT4, '') <> ISNULL(i.IT4, '');

    -- IT5
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IT5',
        d.IT5,
        i.IT5,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IT5, '') <> ISNULL(i.IT5, '');

    -- IT6
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IT6',
        d.IT6,
        i.IT6,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IT6, '') <> ISNULL(i.IT6, '');

    -- IT7
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IT7',
        d.IT7,
        i.IT7,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IT7, '') <> ISNULL(i.IT7, '');

    -- IT8
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IT8',
        d.IT8,
        i.IT8,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IT8, '') <> ISNULL(i.IT8, '');

    -- IT9
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IT9',
        d.IT9,
        i.IT9,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IT9, '') <> ISNULL(i.IT9, '');

    -- IS1
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS1',
        d.IS1,
        i.IS1,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS1, '') <> ISNULL(i.IS1, '');

    -- IS2
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS2',
        d.IS2,
        i.IS2,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS2, '') <> ISNULL(i.IS2, '');

    -- IS3
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS3',
        d.IS3,
        i.IS3,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS3, '') <> ISNULL(i.IS3, '');

    -- IS4
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS4',
        d.IS4,
        i.IS4,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS4, '') <> ISNULL(i.IS4, '');

    -- IS5
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS5',
        d.IS5,
        i.IS5,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS5, '') <> ISNULL(i.IS5, '');

    -- IS6
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS6',
        d.IS6,
        i.IS6,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS6, '') <> ISNULL(i.IS6, '');

    -- IS7
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS7',
        d.IS7,
        i.IS7,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS7, '') <> ISNULL(i.IS7, '');

    -- IS8
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS8',
        d.IS8,
        i.IS8,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS8, '') <> ISNULL(i.IS8, '');

    -- IS9
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS9',
        d.IS9,
        i.IS9,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS9, '') <> ISNULL(i.IS9, '');

    -- IS10
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS10',
        d.IS10,
        i.IS10,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS10, '') <> ISNULL(i.IS10, '');

    -- IS11
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS11',
        d.IS11,
        i.IS11,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS11, '') <> ISNULL(i.IS11, '');

    -- IS12
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS12',
        d.IS12,
        i.IS12,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS12, '') <> ISNULL(i.IS12, '');

    -- IS13
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS13',
        d.IS13,
        i.IS13,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS13, '') <> ISNULL(i.IS13, '');

    -- IS14
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS14',
        d.IS14,
        i.IS14,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS14, '') <> ISNULL(i.IS14, '');

    -- IS15
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS15',
        d.IS15,
        i.IS15,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS15, '') <> ISNULL(i.IS15, '');

    -- IS16
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'IS16',
        d.IS16,
        i.IS16,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.IS16, '') <> ISNULL(i.IS16, '');

    -- SSMTRTAKM
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SSMTRTAKM',
        CONVERT(NVARCHAR(MAX), d.SSMTRTAKM),
        CONVERT(NVARCHAR(MAX), i.SSMTRTAKM),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SSMTRTAKM, 0) <> ISNULL(i.SSMTRTAKM, 0);

    -- SSMTRTAGM
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SSMTRTAGM',
        CONVERT(NVARCHAR(MAX), d.SSMTRTAGM),
        CONVERT(NVARCHAR(MAX), i.SSMTRTAGM),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SSMTRTAGM, 0) <> ISNULL(i.SSMTRTAGM, 0);

    -- SSMSNDAUTO
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SSMSNDAUTO',
        CONVERT(NVARCHAR(MAX), d.SSMSNDAUTO),
        CONVERT(NVARCHAR(MAX), i.SSMSNDAUTO),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SSMSNDAUTO, 0) <> ISNULL(i.SSMSNDAUTO, 0);

    -- SSMTBMON
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SSMTBMON',
        CONVERT(NVARCHAR(MAX), d.SSMTBMON),
        CONVERT(NVARCHAR(MAX), i.SSMTBMON),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SSMTBMON, 0) <> ISNULL(i.SSMTBMON, 0);

    -- SSMDARSAD
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SSMDARSAD',
        CONVERT(NVARCHAR(MAX), d.SSMDARSAD),
        CONVERT(NVARCHAR(MAX), i.SSMDARSAD),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SSMDARSAD, 0) <> ISNULL(i.SSMDARSAD, 0);

    -- HDARKASRTAKHF
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HDARKASRTAKHF',
        d.HDARKASRTAKHF,
        i.HDARKASRTAKHF,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HDARKASRTAKHF, '') <> ISNULL(i.HDARKASRTAKHF, '');

    -- CRT
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'CRT',
        CONVERT(NVARCHAR(MAX), d.CRT, 121),
        CONVERT(NVARCHAR(MAX), i.CRT, 121),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.CRT, '1900-01-01') <> ISNULL(i.CRT, '1900-01-01');

    -- UID
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'UID',
        CONVERT(NVARCHAR(MAX), d.UID),
        CONVERT(NVARCHAR(MAX), i.UID),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.UID, 0) <> ISNULL(i.UID, 0);

    -- PUBLICKEY
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PUBLICKEY',
        d.PUBLICKEY,
        i.PUBLICKEY,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PUBLICKEY, '') <> ISNULL(i.PUBLICKEY, '');

    -- PRIVIATEKEY
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'PRIVIATEKEY',
        d.PRIVIATEKEY,
        i.PRIVIATEKEY,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.PRIVIATEKEY, '') <> ISNULL(i.PRIVIATEKEY, '');

    -- MEMORYID
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'MEMORYID',
        d.MEMORYID,
        i.MEMORYID,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.MEMORYID, '') <> ISNULL(i.MEMORYID, '');

    -- Whether
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'Whether',
        CONVERT(NVARCHAR(MAX), d.Whether),
        CONVERT(NVARCHAR(MAX), i.Whether),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.Whether, 0) <> ISNULL(i.Whether, 0);

    -- MEMORYIDsand
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'MEMORYIDsand',
        d.MEMORYIDsand,
        i.MEMORYIDsand,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.MEMORYIDsand, '') <> ISNULL(i.MEMORYIDsand, '');

    -- Dcertificate
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'Dcertificate',
        d.Dcertificate,
        i.Dcertificate,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.Dcertificate, '') <> ISNULL(i.Dcertificate, '');

    -- HTAHOL
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'HTAHOL',
        CONVERT(NVARCHAR(MAX), CAST(d.HTAHOL AS INT)),
        CONVERT(NVARCHAR(MAX), CAST(i.HTAHOL AS INT)),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.HTAHOL, 0) <> ISNULL(i.HTAHOL, 0);

    -- SMSTYPE
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'SMSTYPE',
        d.SMSTYPE,
        i.SMSTYPE,
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE ISNULL(d.SMSTYPE, '') <> ISNULL(i.SMSTYPE, '');

    -- MOADINA_SCNUM
    INSERT INTO dbo.SAZMAN_Log (UNIVERSITY_CO, ActionType, ColumnName, OldValue, NewValue, 
        ApplicationName, WindowsUsername, ComputerName, IPAddress, LoginName, SessionID, SchemaName, SPID, TransactionID)
    SELECT 
        i.UNIVERSITY_CO, 'UPDATE', 'MOADINA_SCNUM',
        CONVERT(NVARCHAR(MAX), d.MOADINA_SCNUM),
        CONVERT(NVARCHAR(MAX), i.MOADINA_SCNUM),
        @ApplicationName, @WindowsUsername, @ComputerName, @IPAddress, @LoginName, @SessionID, @SchemaName, @SPID, @@TRANCOUNT
    FROM deleted d
    INNER JOIN inserted i ON d.UNIVERSITY_CO = i.UNIVERSITY_CO
    WHERE d.MOADINA_SCNUM <> i.MOADINA_SCNUM;

END;
