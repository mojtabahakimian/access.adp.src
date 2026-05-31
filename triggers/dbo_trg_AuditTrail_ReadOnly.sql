CREATE TRIGGER dbo.trg_AuditTrail_ReadOnly
ON dbo.AuditTrail
INSTEAD OF UPDATE, DELETE
AS
BEGIN
  RAISERROR(N'AuditTrail append-only است؛ ویرایش/حذف مجاز نیست.',16,1);
END;
