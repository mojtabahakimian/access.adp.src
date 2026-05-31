
                        CREATE PROCEDURE [dbo].[sp_CheckReservationTimeout]
                        AS
                        BEGIN
                            SET NOCOUNT ON;
                            SET XACT_ABORT ON;
                            SET LOCK_TIMEOUT 5000;
                            DECLARE @OutputLog TABLE (NUMBER FLOAT);
                            BEGIN TRY
                                BEGIN TRANSACTION;
                                ;WITH TargetReservations AS (
                                    SELECT h.NUMBER, h.TAMIR
                                    FROM dbo.HEAD_LST h
                                    WHERE h.TAG = 20
                                      AND h.TAMIR = 1
                                      AND EXISTS (
                                          SELECT 1
                                          FROM dbo.HEAD_LST_LOG l
                                          WHERE l.NUMBER = h.NUMBER
                                            AND l.TAGG = 20
                                            AND l.UP_DATE < DATEADD(HOUR, -96, GETDATE())
                                      )
                                )
                                UPDATE TargetReservations
                                SET TAMIR = 0
                                OUTPUT inserted.NUMBER INTO @OutputLog(NUMBER);
                                INSERT INTO dbo.HEAD_LST_LOG (UP_DATE, NUMBER, TAGG, RESERVED, UP_USER_NAME, FIELDNAME)
                                SELECT GETDATE(), NUMBER, 20, 0, 'Auto_Job', 'TIMEOUT_CANCELED'
                                FROM @OutputLog;
                                COMMIT TRANSACTION;
                            END TRY
                            BEGIN CATCH
                                IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
                                IF ERROR_NUMBER() = 1222
                                BEGIN
                                    PRINT 'Table is locked by another user. Skipping execution.';
                                END
                                ELSE
                                BEGIN
                                    DECLARE @Err NVARCHAR(MAX) = ERROR_MESSAGE();
                                    RAISERROR(@Err, 16, 1);
                                END
                            END CATCH
                        END
