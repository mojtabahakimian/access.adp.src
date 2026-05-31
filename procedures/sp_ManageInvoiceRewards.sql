CREATE PROCEDURE [dbo].[sp_ManageInvoiceRewards]
								    @InvoiceNumber bigint,
								    @InvoiceTag bigint,
								    @IsRewardSystemActive BIT,
								    @PerformingUserID INT
								  AS
								  BEGIN
								      SET NOCOUNT ON;
								      SET XACT_ABORT ON;
								      
								      DECLARE @CustomerID NVARCHAR(40);
								      DECLARE @InvoiceTotalAmount FLOAT;
								      DECLARE @InvoiceDate BIGINT;
								      DECLARE @CurrentProductCode NVARCHAR(15);
								      DECLARE @TotalProductQuantityInInvoice FLOAT;
								      DECLARE @RewardRuleID INT;
								      DECLARE @RewardType NVARCHAR(50);
								      DECLARE @RewardProductID NVARCHAR(15);
								      DECLARE @RewardQuantity INT;
								      DECLARE @QuantityThreshold INT;
								      DECLARE @RewardDiscountPercentage DECIMAL(5,2);
								      DECLARE @AppliedDiscountAmount FLOAT;
								      DECLARE @NewInvoiceDetailID BIGINT;
								      DECLARE @AnbarIDForReward FLOAT;
								      DECLARE @InvoiceUserName NVARCHAR(40);
								      DECLARE @SourceProductLineID BIGINT;
								      DECLARE @CalculatedRewardQuantity INT; -- مقدار جایزه محاسبه شده
								      
								      BEGIN TRANSACTION;
								      BEGIN TRY
								          -- دریافت اطلاعات فاکتور
								          SELECT
								              @CustomerID = H.CUST_NO,
								              @InvoiceTotalAmount = H.MAS,
								              @InvoiceUserName = H.USER_NAME,
								              @InvoiceDate = H.DATE_N
								          FROM dbo.HEAD_LST AS H
								          WHERE H.NUMBER = @InvoiceNumber AND H.TAG = @InvoiceTag;
								  
								          IF @CustomerID IS NULL
								          BEGIN
								              RAISERROR('فاکتور با شماره و تگ مشخص شده یافت نشد.', 16, 1);
								              RETURN;
								          END;
								  
								          -- حذف جوایز قبلی
								          DECLARE previous_rewards_cursor CURSOR LOCAL FAST_FORWARD FOR
								          SELECT IL.CODE, IL.MEGH, IL.ANBAR
								          FROM dbo.INVO_LST AS IL
								          WHERE IL.NUMBER = @InvoiceNumber
								              AND IL.TAG = @InvoiceTag
								              AND ISNULL(IL.JAY, 0) > 0;
								  
								          OPEN previous_rewards_cursor;
								          FETCH NEXT FROM previous_rewards_cursor INTO @RewardProductID, @RewardQuantity, @AnbarIDForReward;
								          WHILE @@FETCH_STATUS = 0
								          BEGIN
								              IF @RewardProductID IS NOT NULL AND @RewardQuantity IS NOT NULL AND @AnbarIDForReward IS NOT NULL
								              BEGIN
								                  UPDATE dbo.STUF_STK
								                  SET MOGODI_A = MOGODI_A + @RewardQuantity
								                  WHERE CODE = @RewardProductID AND ANBAR = @AnbarIDForReward;
								              END
								              FETCH NEXT FROM previous_rewards_cursor INTO @RewardProductID, @RewardQuantity, @AnbarIDForReward;
								          END;
								          CLOSE previous_rewards_cursor;
								          DEALLOCATE previous_rewards_cursor;
								  
								          -- حذف سطرهای جایزه قبلی
								          DELETE FROM dbo.INVO_LST
								          WHERE NUMBER = @InvoiceNumber
								              AND TAG = @InvoiceTag
								              AND ISNULL(JAY, 0) > 0;
								  
								          DELETE FROM dbo.InvoiceRewards
								          WHERE InvoiceNumber = @InvoiceNumber AND InvoiceTag = @InvoiceTag;
								  
								          -- اعمال جوایز جدید
								          IF @IsRewardSystemActive = 1
								          BEGIN
								              DECLARE product_cursor CURSOR LOCAL FAST_FORWARD FOR
								              SELECT IL.CODE, IL.ANBAR
								              FROM dbo.INVO_LST AS IL
								              WHERE IL.NUMBER = @InvoiceNumber
								                  AND IL.TAG = @InvoiceTag
								                  AND ISNULL(IL.JAY, 0) = 0
								              GROUP BY IL.CODE, IL.ANBAR;
								  
								              OPEN product_cursor;
								              FETCH NEXT FROM product_cursor INTO @CurrentProductCode, @AnbarIDForReward;
								  
								              WHILE @@FETCH_STATUS = 0
								              BEGIN
								                  -- محاسبه مجموع مقدار کالا در فاکتور
								                  SELECT @TotalProductQuantityInInvoice = ISNULL(SUM(IL.MEGHk), 0)
								                  FROM dbo.INVO_LST AS IL
								                  WHERE IL.NUMBER = @InvoiceNumber
								                      AND IL.TAG = @InvoiceTag
								                      AND IL.CODE = @CurrentProductCode
								                      AND IL.ANBAR = @AnbarIDForReward
								                      AND ISNULL(IL.JAY, 0) = 0;
								  
								                  -- دریافت شناسه اولین ردیف کالای اصلی
								                  SELECT TOP 1 @SourceProductLineID = IL.id
								                  FROM dbo.INVO_LST AS IL
								                  WHERE IL.NUMBER = @InvoiceNumber
								                      AND IL.TAG = @InvoiceTag
								                      AND IL.CODE = @CurrentProductCode
								                      AND IL.ANBAR = @AnbarIDForReward
								                      AND ISNULL(IL.JAY, 0) = 0
								                  ORDER BY IL.id ASC;
								  
								                  -- پردازش تمام قوانین جایزه قابل اعمال
								                  DECLARE reward_rules_cursor CURSOR LOCAL FAST_FORWARD FOR
								                  SELECT 
								                      RR.RuleID, 
								                      RR.Reward_Type, 
								                      RR.Reward_ProductID, 
								                      RR.Reward_Quantity, 
								                      RR.Quantity_Threshold,
								                      RR.Reward_Discount_Percentage
								                  FROM dbo.RewardRules AS RR
								                  WHERE RR.ProductID_Target = @CurrentProductCode
								                      AND RR.IsActive = 1
								                      AND (RR.StartDate IS NULL OR RR.StartDate <= @InvoiceDate)
								                      AND (RR.EndDate IS NULL OR RR.EndDate >= @InvoiceDate)
								                      AND @TotalProductQuantityInInvoice >= RR.Quantity_Threshold
								                  ORDER BY RR.Quantity_Threshold DESC;
								  
								                  OPEN reward_rules_cursor;
								                  FETCH NEXT FROM reward_rules_cursor INTO 
								                      @RewardRuleID, @RewardType, @RewardProductID, 
								                      @RewardQuantity, @QuantityThreshold, @RewardDiscountPercentage;
								  
								                  WHILE @@FETCH_STATUS = 0 AND @SourceProductLineID IS NOT NULL
								                  BEGIN
								                      -- محاسبه مقدار جایزه بر اساس تعداد دفعات برآورده شدن threshold
								                      SET @CalculatedRewardQuantity = 
								                          (CAST(@TotalProductQuantityInInvoice AS INT) / @QuantityThreshold) * @RewardQuantity;
								  
								                      IF @RewardType = 'Product' AND @RewardProductID IS NOT NULL AND @CalculatedRewardQuantity > 0
								                      BEGIN
								                          -- Ensure the product exists in the warehouse (STUF_FSK) to prevent FK violation
								                          IF NOT EXISTS (SELECT 1 FROM dbo.STUF_FSK WHERE CODE = @RewardProductID AND ANBAR = @AnbarIDForReward)
								                          BEGIN
								                               INSERT INTO dbo.STUF_FSK (CODE, ANBAR, MOGODI_A, FI_A)
								                               VALUES (@RewardProductID, @AnbarIDForReward, 0, 0);
								                          END

								                          -- درج ردیف جایزه در INVO_LST
								                          INSERT INTO dbo.INVO_LST (
								                              NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, 
								                              MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, 
								                              ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, IMBAA, TOTALARZ, 
								                              VISITOR, TKHN, JAY, JAYO, CRT, UID
								                          )
								                          SELECT
								                              @InvoiceNumber, 
								                              @InvoiceTag, 
								                              @AnbarIDForReward,
								                              (SELECT ISNULL(MAX(RADIF), 0) + 1 FROM dbo.INVO_LST 
								                               WHERE NUMBER = @InvoiceNumber AND TAG = @InvoiceTag),
								                              @RewardProductID, 
								                              CAST(@CalculatedRewardQuantity AS FLOAT), -- مقدار محاسبه شده
								                              CAST(@CalculatedRewardQuantity AS FLOAT),
								                              0, NULL, 1, CAST(@CalculatedRewardQuantity AS FLOAT), 0, NULL, 0, NULL, NULL, NULL, NULL,
								                              (SELECT
								                                  CASE
								                                      WHEN ISNUMERIC(SDEF.VAHED) = 1
								                                      THEN CONVERT(FLOAT, SDEF.VAHED)
								                                      ELSE NULL
								                                  END
								                              FROM dbo.STUF_DEF SDEF WHERE SDEF.CODE = @RewardProductID),
								                              100, CAST(@CalculatedRewardQuantity AS FLOAT), NULL, 0, 0, 0, @InvoiceUserName, 0,
								                              @SourceProductLineID, 
								                              NULL, GETDATE(), @PerformingUserID;
								  
								                          SELECT @NewInvoiceDetailID = SCOPE_IDENTITY();
								  
								                          -- کسر از موجودی انبار
								                          UPDATE SF
								                          SET MOGODI_A = SF.MOGODI_A - @CalculatedRewardQuantity
								                          FROM dbo.STUF_STK AS SF
								                          WHERE SF.CODE = @RewardProductID AND SF.ANBAR = @AnbarIDForReward;
								  
								                          -- ثبت در جدول InvoiceRewards
								                          INSERT INTO dbo.InvoiceRewards (
								                              InvoiceNumber, InvoiceTag, CustomerID, RewardRuleID,
								                              ProductCode_Earned, Quantity_Earned, Reward_Given_Type,
								                              Reward_Given_ProductCode, Reward_Given_Quantity, Reward_Given_Discount_Amount,
								                              RewardDate, RecordedBy_UserID, CRT, UID
								                          )
								                          VALUES (
								                              @InvoiceNumber, @InvoiceTag, @CustomerID, @RewardRuleID,
								                              @CurrentProductCode, @TotalProductQuantityInInvoice, @RewardType,
								                              @RewardProductID, @CalculatedRewardQuantity, 0,
								                              @InvoiceDate, @PerformingUserID, GETDATE(), @PerformingUserID
								                          );
								                      END
								                      ELSE IF @RewardType = 'Discount'
								                      BEGIN
								                          SET @AppliedDiscountAmount = 0;
								                          -- منطق تخفیف در صورت نیاز
								                      END;
								  
								                      FETCH NEXT FROM reward_rules_cursor INTO 
								                          @RewardRuleID, @RewardType, @RewardProductID, 
								                          @RewardQuantity, @QuantityThreshold, @RewardDiscountPercentage;
								                  END;
								  
								                  CLOSE reward_rules_cursor;
								                  DEALLOCATE reward_rules_cursor;
								  
								                  FETCH NEXT FROM product_cursor INTO @CurrentProductCode, @AnbarIDForReward;
								              END;
								              CLOSE product_cursor;
								              DEALLOCATE product_cursor;
								          END;
								  
								          COMMIT TRANSACTION;
								          SELECT 'Reward management process completed successfully.' AS Result;
								  
								      END TRY
								      BEGIN CATCH
								          DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
								          DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
								          DECLARE @ErrorState INT = ERROR_STATE();
								  
								          IF @@TRANCOUNT > 0
								              ROLLBACK TRANSACTION;
								  
								          RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
								          RETURN;
								      END CATCH;
								  END
