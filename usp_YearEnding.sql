USE [UniteamBase]
GO
/****** Object:  StoredProcedure [dbo].[usp_YearEnding]    Script Date: 6/3/2019 12:36:02 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Yuan Weixin>
-- Create date: <2012-12-25>
-- Description:	<Procedure to do the year ending logic>
-- =============================================
ALTER PROCEDURE [dbo].[usp_YearEnding]
	@CompanyID      int,
	@UserID         int,
	@YearID         int,
	@YeResNomID     int,
	@YeClSaleNomID  int,
	@YeClPurchNomID int,
	@DayBook        int output
AS
BEGIN TRY
    BEGIN TRANSACTION
    
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
    
    --Company related variables
	Declare @Company_BasicCurrID     nvarchar(6)
    Declare @Company_FilePostPeriod  tinyint    
    Declare @Company_PostingYear     tinyint     
    Declare @Company_StartPostPeriod tinyint 
    Select @Company_FilePostPeriod = [FilePostPeriodID],
           @Company_PostingYear = [PostingYear],
           @Company_StartPostPeriod =[StartPostPeriod],
           @Company_BasicCurrID = [BasicCurrID]
    From [dbo].[Company]
    Where [CompanyID] = @CompanyID
    
    Declare @Local_NewYear       int
    Declare @Local_NewPeriod     int
    Declare @Local_NewDate       date
    Declare @Local_FromPeriod    int
    Declare @Local_ToPeriod      int
	Declare @Local_LastNomID     int
	Declare @Local_LastServiceID int
	Declare @Local_LastFileID    int
	Declare @Local_LastActID     int
	Declare @Local_LastBranchID  int
	Declare @Local_LastNomTypeID nchar(1)
	Declare @Local_LastAmount     numeric(14,2)
	Declare @Local_LastCurrAmount numeric(14,2)
	Declare @Local_LastFileFinNomID int
	Declare @Local_LastFileFinInd   nchar(1)

    Declare @Local_YearResult    numeric(14,2)    
    Declare @Local_SysDateTime   datetime      
    
	Declare @CounterID int
	Declare @ProvisionalSeqRequired bit
	Declare @TransNo Int

    Set @Local_NewYear     = @YearID + 1
    Set @Local_NewPeriod   = (@Local_NewYear % 100) * 100
    Set @Local_NewDate     = CONVERT(varchar(4), @YearID) + (STUFF('00', 1, len(@Company_StartPostPeriod),'') + CONVERT(varchar(2), @Company_StartPostPeriod)) + '01'
    Set @Local_FromPeriod  = (@YearID % 100) * 100
    Set @Local_ToPeriod    = @Local_FromPeriod + 13    
    Set @Local_YearResult  = 0.00    
    Set @Local_SysDateTime = SYSDATETIME() 

	DECLARE @Temp_FP TABLE (ID INT IDENTITY(1,1), RecID INT, NomID INT, ServiceID INT, FileID INT, ActID INT, BranchID INT, Amount numeric(14,2), CurrAmount numeric(14,2), FileFinNomID int, FileFinInd nchar(1))

  
    
	--#1: 1st pass, excluding expected profit in balance sheet
	Delete @Temp_FP
	Insert into @Temp_FP (RecID, NomID, ServiceID, ActID, BranchID, Amount, CurrAmount, FileFinNomID, FileFinInd)
	Select f.RecID,
	       [NomID] = (Case when f.[TransType] = 'A' then (Case when l.[NomTypeID] = 'A' then @YeClPurchNomID
															   When l.[NomTypeID] = 'L' then @YeClSaleNomID
															   else f.[NomID] End)
						   when f.[FileFinInd] = '3' then (Case when l.[NomTypeID] = 'A' or l.[NomTypeID] = 'C' then @YeClPurchNomID
																else @YeClSaleNomID End)
						   else f.[NomID] End),
		   [ServiceID] = f.[ServiceID],
		   [ActID] = f.[ActID],
		   [BranchID] = f.[BranchID],
		   [Amount] = Case when f.Amount is null then 0.00 else f.Amount end,
		   [CurrAmount] = Case when f.CurrAmount is null then 0.00 else f.CurrAmount end,
		   f.FileFinNomID,
		   f.FileFinInd
	From dbo.FinancePosting f
	Left join dbo.Ledger l on l.CompanyID = f.CompanyID and l.NomID = f.NomID
	Left join dbo.[Service] s on s.ServiceID = f.ServiceID
	Left join dbo.Activity a on a.ActivityID = f.ActID
	Left join dbo.Branch b on b.BranchID = f.BranchID
	Where f.CompanyID = @CompanyID AND 
	      f.Period >= @Local_FromPeriod AND 
		  f.Period <= @Local_ToPeriod AND
		  (f.TransType <> 'S' OR l.NomTypeID IN ('I', 'C'))
	Order by l.NomTypeID, f.NomID, f.FileFinNomID, s.ServiceCode, a.ActivityCode, b.BranchCode, f.ItemNo

	Declare @Temp_ID INT
	Declare @Temp_RecID INT
	Select @Temp_ID = MIN(ID) From @Temp_FP
	WHILE @Temp_ID IS NOT NULL
	Begin
		Select @Temp_RecID = RecID, 
		       @Local_LastNomID = NomID, 
			   @Local_LastServiceID = ServiceID, 
			   @Local_LastActID = ActID, 
			   @Local_LastBranchID = BranchID,
			   @Local_LastFileFinNomID = FileFinNomID,
			   @Local_LastFileFinInd = FileFinInd
		From @Temp_FP Where ID = @Temp_ID
	
	    Select @Local_LastNomTypeID = NomTypeID
		From dbo.Ledger where CompanyID = @CompanyID AND NomID = @Local_LastNomID

		Set @Local_LastAmount = 0.00
		Set @Local_LastCurrAmount = 0.00
		Select @Local_LastAmount = Sum(Amount), @Local_LastCurrAmount = Sum(CurrAmount)
		From @Temp_FP
		Where NomID = @Local_LastNomID AND 
		      ((@Local_LastServiceID IS NULL AND ServiceID IS NULL) OR ServiceID = @Local_LastServiceID) AND 
			  ((@Local_LastActID IS NULL AND ActID IS NULL) OR ActID = @Local_LastActID) AND 
			  ((@Local_LastBranchID IS NULL AND BranchID IS NULL) OR BranchID = @Local_LastBranchID)

		If (@Local_LastNomTypeID IN ('A', 'C') OR (@Local_LastFileFinNomID IS NOT NULL AND @Local_LastFileFinInd IS NULL)) AND @Local_LastAmount <> 0 AND @Local_LastAmount IS NOT NULL
		Begin
		    --use the counter system to get the DayBook and TransNo
			If @DayBook IS NULL OR @DayBook = 0
			Begin
				--A: Get CounterID
				Select @CounterID = [AdvancedCounterID], @ProvisionalSeqRequired = [ProvisionalSequenceRequired]
				From [dbo].[AdvancedCounterDefinitions]
				Where [Name] = 'DayBook'
	    
				--B: Allocate the seq
				If @ProvisionalSeqRequired = 1
					EXEC [dbo].[usp_AllocateProvisionalAdvancedSequence] @CounterID, @CompanyID, @DayBook OUTPUT	
		
				EXEC [dbo].[usp_AllocateAdvancedSequence] @CounterID, @CompanyID, @DayBook OUTPUT

				--TransNo				
				EXEC [dbo].[usp_AllocateSimpleSequence]	@SimpleCounterID = 24, @Name = N'TransNo', @CompanyID = null, @SequenceNumber = @TransNo OUTPUT
			End 

			Insert into [dbo].[FinancePosting]
			(	[CompanyID],
				[ItemNo],
				[BookID],
				[RollBackInd],
				[TransType],
				[NomID],
				[Period],
				[PostDate],
				[ServiceID],
				[ActID],
				[ItemText],
				[CurrID],
				[Amount],
				[CurrAmount],
				[VATCode],
				[FileID],
				[UserID],
				[BranchID],
				[DayBook],
				[FileFinNomID],
				[FileFinInd],
				[OrderID],
				[CustID],
				[TransNo],
				[LinkID],
				[LinkID2])
			Select [CompanyID],
				   [ItemNo] = 1,
				   [BookID],
				   [RollBackInd] = '0',
				   [TransType],
				   [NomID],
				   [Period] = @Local_NewPeriod,
				   [PostDate] = @Local_NewDate,
				   [ServiceID],
				   [ActID],
				   [ItemText] = 'PRIMO',
				   [CurrID],
				   [Amount] = @Local_LastAmount,
				   [CurrAmount] = @Local_LastCurrAmount,
				   [VATCode] = [VATCode],
				   [FileID] = (Case when @Local_LastNomID = @YeClPurchNomID or @Local_LastNomID = @YeClSaleNomID then null
									else [FileID] End),
				   [UserID] = @UserID,
				   [BranchID],
				   [DayBook] = @DayBook,
				   [FileFinNomID],
				   [FileFinInd],
				   [OrderID] = null,
				   [CustID] = (Case when @Local_LastNomID = @YeClPurchNomID or @Local_LastNomID = @YeClSaleNomID then null
									when [FileFinNomID] is null then null else [CustID] End),
				   [TransNo] = @TransNo,
				   [LinkID] = null,
				   [LinkID2] = null				
			From dbo.FinancePosting
			Where RecID = @Temp_RecID
		End
		Else
			Set @Local_YearResult = @Local_YearResult + @Local_LastAmount

		Select @Temp_ID = MIN(ID) From @Temp_FP Where ID > @Temp_ID
	End

	    
	--#2: 2st pass, eliminating expected profit in balance sheet if sum is zero (per. Svc/file/Branch/Act)
	Delete @Temp_FP
	Insert into @Temp_FP (RecID, NomID, ServiceID, FileID, ActID, BranchID, Amount, CurrAmount, FileFinNomID, FileFinInd)
	Select f.RecID,
	       [NomID] = f.[NomID],
		   [ServiceID] = f.[ServiceID],
		   [FileID] = f.[FileID],
		   [ActID] = f.[ActID],
		   [BranchID] = f.[BranchID],
		   [Amount] = Case when f.Amount is null then 0.00 else f.Amount end,
		   [CurrAmount] = Case when f.CurrAmount is null then 0.00 else f.CurrAmount end,
		   f.FileFinNomID,
		   f.FileFinInd
	From dbo.FinancePosting f
	Left join dbo.Ledger l on l.CompanyID = f.CompanyID and l.NomID = f.NomID
	Left join dbo.[Service] s on s.ServiceID = f.ServiceID
	Left join dbo.[File] m on m.FileID = f.FileID
	Left join dbo.Activity a on a.ActivityID = f.ActID
	Left join dbo.Branch b on b.BranchID = f.BranchID
	Where f.CompanyID = @CompanyID AND 
	      f.Period >= @Local_FromPeriod AND 
		  f.Period <= @Local_ToPeriod AND
		  (f.TransType = 'S' AND l.NomTypeID IN ('A', 'L'))
	Order by l.NomTypeID, f.NomID, s.ServiceCode, m.FileNumberID, a.ActivityCode, b.BranchCode

	Select @Temp_ID = MIN(ID) From @Temp_FP
	WHILE @Temp_ID IS NOT NULL
	Begin
		Select @Temp_RecID = RecID, 
		       @Local_LastNomID = NomID, 
			   @Local_LastServiceID = ServiceID,
			   @Local_LastFileID = FileID, 
			   @Local_LastActID = ActID, 
			   @Local_LastBranchID = BranchID
		From @Temp_FP Where ID = @Temp_ID
	
	    Select @Local_LastNomTypeID = NomTypeID
		From dbo.Ledger where CompanyID = @CompanyID AND NomID = @Local_LastNomID
	    
		Set @Local_LastAmount = 0.00
		Set @Local_LastCurrAmount = 0.00
		Select @Local_LastAmount = Sum(Amount), @Local_LastCurrAmount = Sum(CurrAmount)
		From @Temp_FP
		Where NomID = @Local_LastNomID AND 
		      ((@Local_LastServiceID IS NULL AND ServiceID IS NULL) OR ServiceID = @Local_LastServiceID) AND 
			  ((@Local_LastFileID IS NULL AND FileID IS NULL) OR FileID = @Local_LastFileID) AND 
			  ((@Local_LastActID IS NULL AND ActID IS NULL) OR ActID = @Local_LastActID) AND 
			  ((@Local_LastBranchID IS NULL AND BranchID IS NULL) OR BranchID = @Local_LastBranchID)

		If @Local_LastAmount <> 0 AND @Local_LastAmount IS NOT NULL
		Begin
		    --use the counter system to get the DayBook and TransNo
			If @DayBook IS NULL OR @DayBook = 0
			Begin
				--A: Get CounterID
				Select @CounterID = [AdvancedCounterID], @ProvisionalSeqRequired = [ProvisionalSequenceRequired]
				From [dbo].[AdvancedCounterDefinitions]
				Where [Name] = 'DayBook'
	    
				--B: Allocate the seq
				If @ProvisionalSeqRequired = 1
					EXEC [dbo].[usp_AllocateProvisionalAdvancedSequence] @CounterID, @CompanyID, @DayBook OUTPUT	
		
				EXEC [dbo].[usp_AllocateAdvancedSequence] @CounterID, @CompanyID, @DayBook OUTPUT

				--TransNo
				EXEC [dbo].[usp_AllocateSimpleSequence]	@SimpleCounterID = 24, @Name = N'TransNo', @CompanyID = null, @SequenceNumber = @TransNo OUTPUT
			End 

			Insert into [dbo].[FinancePosting]
			(	[CompanyID],
				[ItemNo],
				[BookID],
				[RollBackInd],
				[TransType],
				[NomID],
				[Period],
				[PostDate],
				[ServiceID],
				[ActID],
				[ItemText],
				[CurrID],
				[Amount],
				[CurrAmount],
				[VATCode],
				[FileID],
				[UserID],
				[BranchID],
				[DayBook],
				[FileFinNomID],
				[FileFinInd],
				[OrderID],
				[CustID],
				[TransNo],
				[LinkID],
				[LinkID2])
			Select [CompanyID],
				   [ItemNo] = 1,
				   [BookID],
				   [RollBackInd] = '0',
				   [TransType],
				   [NomID],
				   [Period] = @Local_NewPeriod,
				   [PostDate] = @Local_NewDate,
				   [ServiceID],
				   [ActID],
				   [ItemText] = 'PRIMO',
				   [CurrID],
				   [Amount] = @Local_LastAmount,
				   [CurrAmount] = @Local_LastCurrAmount,
				   [VATCode] = [VATCode],
				   [FileID],
				   [UserID] = @UserID,
				   [BranchID],
				   [DayBook] = @DayBook,
				   [FileFinNomID],
				   [FileFinInd],
				   [OrderID] = null,
				   [CustID],
				   [TransNo] = @TransNo,
				   [LinkID] = null,
				   [LinkID2] = null				
			From dbo.FinancePosting
			Where RecID = @Temp_RecID
		End
		Else
			Set @Local_YearResult = @Local_YearResult + @Local_LastAmount

		Select @Temp_ID = MIN(ID) From @Temp_FP Where ID > @Temp_ID
	End

	--#3: Post year end result
	Begin
        --#4-A: Insert the year end result to the FinancePosting table
		Insert into [dbo].[FinancePosting]
		(
			[CompanyID],
			[ItemNo],
			[BookID],
			[RollBackInd],
			[TransType],
			[NomID],
			[Period],
			[PostDate],
			[ServiceID],
			[ActID],
			[ItemText],
			[CurrID],
			[Amount],
			[CurrAmount],
			[VATCode],
			[FileID],
			[UserID],
			[BranchID],
			[DayBook],
			[FileFinNomID],[FileFinInd],[OrderID],[CustID],[TransNo],[LinkID],[LinkID2]
		)
		Values
		(
			@CompanyID,
			1,
			null,
			'0',
			'B',
			@YeResNomID,
			@Local_NewPeriod,
			@Local_NewDate,
			null,
			null,
			'PRIMO',
			@Company_BasicCurrID,
			@Local_YearResult,
			@Local_YearResult,
			'0',
			null,
			@UserID,
			null,
			@DayBook,
			null,null,null,null,@TransNo,null,null
		)
		
		--#4-B: Update the Company
		Update [dbo].[Company]
		Set [IsYearEndDone] = 1
		Where [CompanyID] = @CompanyID
        
        --#4-C: Insert the year end period to the PerEnding table
        Begin
			Declare @closeDate date
			Declare @closeTime int
			Set @closeDate = Convert(varchar(10),@Local_SysDateTime, 112)
			Set @closeTime = DATEPART(hour, @Local_SysDateTime)*100 + DATEPART(minute, @Local_SysDateTime)
			
			Insert into [dbo].[PerEnding]
			([CompanyID],[PeriodID],[UserID],[CloseDate],[CloseTime])
			values
			(@CompanyID, @Local_NewYear*100, @UserID, @closeDate, @closeTime)
        End
		
		--#4-D: Insert the year end to the YearEnding table
		Insert into [dbo].[YearEnding]
		([CompanyID], [PostingYear], [StartPostPeriod])
		Values
		(@CompanyID, @Local_NewYear % 100, @Company_StartPostPeriod)
	End
	
	--#4: Commit the transaction
    COMMIT TRANSACTION
END TRY
BEGIN CATCH
    ROLLBACK TRANSACTION 
END CATCH
