
/*-------------------------------------------------------------
  3) ساخت تریگر
-------------------------------------------------------------*/
CREATE TRIGGER [dbo].[trg_STUF_FSK_Audit]
ON [dbo].[STUF_FSK]
AFTER INSERT, UPDATE, DELETE
AS
SET NOCOUNT ON;

DECLARE @host nvarchar(128) = HOST_NAME();

/* ----------------- INSERT ----------------- */
INSERT INTO dbo.STUF_FSK_Audit
(
  AuditHost, Operation,
  CODE, ANBAR, IDD,
  New_MOGODI_A, New_FI_A, New_MABL_A, New_MANDAH_A, New_VAZ, New_POSITION,
  New_B_SEF, New_N_SEF, New_MIN_M, New_MAX_M, New_CRT, New_UID
)
SELECT
  @host, 'I',
  i.CODE, i.ANBAR, i.IDD,
  i.MOGODI_A, i.FI_A, i.MABL_A, i.MANDAH_A, i.VAZ, i.POSITION,
  i.B_SEF, i.N_SEF, i.MIN_M, i.MAX_M, i.CRT, i.UID
FROM inserted i
LEFT JOIN deleted d
       ON  d.CODE  = i.CODE
       AND d.ANBAR = i.ANBAR
WHERE d.CODE IS NULL;    -- واقعاً رکورد تازه است

/* ----------------- DELETE ----------------- */
INSERT INTO dbo.STUF_FSK_Audit
(
  AuditHost, Operation,
  CODE, ANBAR, IDD,
  Old_MOGODI_A, Old_FI_A, Old_MABL_A, Old_MANDAH_A, Old_VAZ, Old_POSITION,
  Old_B_SEF, Old_N_SEF, Old_MIN_M, Old_MAX_M, Old_CRT, Old_UID
)
SELECT
  @host, 'D',
  d.CODE, d.ANBAR, d.IDD,
  d.MOGODI_A, d.FI_A, d.MABL_A, d.MANDAH_A, d.VAZ, d.POSITION,
  d.B_SEF, d.N_SEF, d.MIN_M, d.MAX_M, d.CRT, d.UID
FROM deleted d
LEFT JOIN inserted i
       ON  i.CODE  = d.CODE
       AND i.ANBAR = d.ANBAR
WHERE i.CODE IS NULL;    -- واقعاً حذف شده است

/* ----------------- UPDATE ----------------- */
INSERT INTO dbo.STUF_FSK_Audit
(
  AuditHost, Operation,
  CODE, ANBAR, IDD,
  Old_MOGODI_A, New_MOGODI_A,
  Old_FI_A,     New_FI_A,
  Old_MABL_A,   New_MABL_A,
  Old_MANDAH_A, New_MANDAH_A,
  Old_VAZ,      New_VAZ,
  Old_POSITION, New_POSITION,
  Old_B_SEF,    New_B_SEF,
  Old_N_SEF,    New_N_SEF,
  Old_MIN_M,    New_MIN_M,
  Old_MAX_M,    New_MAX_M,
  Old_CRT,      New_CRT,
  Old_UID,      New_UID
)
SELECT
  @host, 'U',
  d.CODE, d.ANBAR, d.IDD,
  d.MOGODI_A, i.MOGODI_A,
  d.FI_A,     i.FI_A,
  d.MABL_A,   i.MABL_A,
  d.MANDAH_A, i.MANDAH_A,
  d.VAZ,      i.VAZ,
  d.POSITION, i.POSITION,
  d.B_SEF,    i.B_SEF,
  d.N_SEF,    i.N_SEF,
  d.MIN_M,    i.MIN_M,
  d.MAX_M,    i.MAX_M,
  d.CRT,      i.CRT,
  d.UID,      i.UID
FROM deleted d
JOIN inserted i
     ON  i.CODE  = d.CODE
     AND i.ANBAR = d.ANBAR;
