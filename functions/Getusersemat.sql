CREATE FUNCTION [dbo].[Getusersemat]
									(
									    @usid INT,
									    @fld NVARCHAR(50)
									)
									RETURNS NVARCHAR(100)
									AS
									BEGIN
									    DECLARE @ret NVARCHAR(100)
									
									    SELECT @ret = 
									        CASE 
									            WHEN ISNULL(
									                CASE @fld
									                    WHEN 'FFR_FROOSHTX' THEN FFR_FROOSHTX
									                    WHEN 'FFR_HESABTX'  THEN FFR_HESABTX
									                    WHEN 'FFR_MODIRTX'  THEN FFR_MODIRTX
									                END, ''
									            ) <> '' THEN 
									                CASE @fld
									                    WHEN 'FFR_FROOSHTX' THEN FFR_FROOSHTX
									                    WHEN 'FFR_HESABTX'  THEN FFR_HESABTX
									                    WHEN 'FFR_MODIRTX'  THEN FFR_MODIRTX
									                END
									            ELSE 
									                CASE @fld
									                    WHEN 'FFR_FROOSHTX' THEN N'فروش'
									                    WHEN 'FFR_HESABTX'  THEN N'حسابداري'
									                    WHEN 'FFR_MODIRTX'  THEN N'مدير عامل'
									                    ELSE N''
									                END
									        END
									    FROM SIGN
									    WHERE USERCO = @usid
									
									    RETURN ISNULL(@ret, N'')
									END
