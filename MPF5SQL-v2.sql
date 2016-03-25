-- output: K:\MRO\Archive\yyyy\yyyy-mm\yyyy-mm-dd\Pricing\Pricing-HFV-Loans.csv
--         TPODSLoc
--		polypaths Input	---==== create pf file
--		Instrument_MBS_MPF ==== to MiddleWare
--		Portfolio_Contents_MBS_MPF ==== to MiddleWare


-- qryMissingHFSPrices ======= cmdCheckMissingHFSPrices_Click
SELECT Cusip, LoanNumber, Price
FROM tblHFSLoansEOD
WHERE (Price Is Null) OR (Price = 0);

-- qryDeleteMPFandForwardsfromInstrumenttblMiddleware ===== cmdDeleteMiddleware_Click
DELETE FROM Instrument_MBS_MPF
WHERE (((Instrument_MBS_MPF.CUSIP) Like 'AA%' Or (Instrument_MBS_MPF.CUSIP) Like 'GL%' Or (Instrument_MBS_MPF.CUSIP) Like 'MS%' Or (Instrument_MBS_MPF.CUSIP) Like 'SS%' Or (Instrument_MBS_MPF.CUSIP) Like 'DC%' Or (Instrument_MBS_MPF.CUSIP) Like 'HC%' Or (Instrument_MBS_MPF.CUSIP) Like 'MA%') AND ((Instrument_MBS_MPF.BondCUSIP)='MPF' Or (Instrument_MBS_MPF.BondCUSIP)='MPFForward' Or (Instrument_MBS_MPF.BondCUSIP)='MPFAct'))

-- qryDeleteDataFromPortfolioContentsMiddleware ========= cmdDeleteMiddleware_Click()
DELETE FROM Portfolio_Contents_MBS_MPF

-- qryHFSLoanPricesToTPODS ========= cmdExportHFSToTPODS_Click
SELECT FORMAT(Date(),"yyyymmdd") AS PricingDate, tblHFSLoansEOD.LoanNumber, tblHFSLoansEOD.Price
FROM tblHFSLoansEOD;


--------------- 20140820

-- ===== cmdInsertMiddleware_Click  ------------- begin
-- qryAppendMPFSecuritiesToInstrumentTableMiddleware
INSERT INTO Instrument_MBS_MPF ( CUSIP, Active_q, Agency, Delay, Wac, Coup, Balloon, Owam, IF, PF, FA1, Const1, Mult1, Rate1, RF1, Floor1, Cap1, PF1, PC1, AB1, LookBack1, Int_Coup_Mult, PP_Coup_Mult, Sum_Floor, Sum_Cap, Servicing_Model, Loss_Model, Sched_Cap_q, BondCUSIP )
SELECT CUSIP, [Active?], Agency, Delay, AggWac, 
	AggCoup, Ballon, AggOWAM, IF, PF, FA, Const1, Mult1, Rate, RF, Floor, Cap, PF1 AS PF1, PC, AB, LookBack, [Int Coup Mult], [PP Coup Mult], [Sum Floor], [Sum Cap], [Servicing Model], [Loss Model], [Sched Cap?], Portfolio
FROM tblSecuritiesPalmsSource;

---- qryAppendMPFMarkToModelToInstrumentTableMiddleware
INSERT INTO Instrument_MBS_MPF ( CUSIP, Active_q, Agency, Delay, Wac, Coup, Balloon, Owam, IF, PF, FA1, Const1, Mult1, Rate1, RF1, Floor1, Cap1, PF1, PC1, AB1, LookBack1, Int_Coup_Mult, PP_Coup_Mult, Sum_Floor, Sum_Cap, Servicing_Model, Loss_Model, Sched_Cap_q, BondCUSIP )
SELECT tblMarkToModelLoans.CUSIP, tblMarkToModelLoans.[Active?], tblMarkToModelLoans.Agency, tblMarkToModelLoans.Delay, tblMarkToModelLoans.AggWac, tblMarkToModelLoans.AggCoup, tblMarkToModelLoans.Ballon, tblMarkToModelLoans.AggOWAM, tblMarkToModelLoans.IF, tblMarkToModelLoans.PF, tblMarkToModelLoans.FA, tblMarkToModelLoans.Const1, tblMarkToModelLoans.Mult1, tblMarkToModelLoans.Rate, tblMarkToModelLoans.RF, tblMarkToModelLoans.Floor, tblMarkToModelLoans.Cap, tblMarkToModelLoans.PF1 AS PF1, tblMarkToModelLoans.PC, tblMarkToModelLoans.AB, tblMarkToModelLoans.LookBack, tblMarkToModelLoans.[Int Coup Mult], tblMarkToModelLoans.[PP Coup Mult], tblMarkToModelLoans.[Sum Floor], tblMarkToModelLoans.[Sum Cap], tblMarkToModelLoans.[Servicing Model], tblMarkToModelLoans.[Loss Model], tblMarkToModelLoans.[Sched Cap?], tblMarkToModelLoans.Portfolio
FROM tblMarkToModelLoans;

--- qryAppendMPFSecuritiesToPortfolioContentsTableMiddleware
INSERT INTO Portfolio_Contents_MBS_MPF ( Portfolio, IntexID, CUSIP, Account_Class, Account_Type, Sub_Account_Type, H1, H2, Notional, Mult, Factor, Add_Accrued_q, PV, Swap, Wac, Coup, Wam, Age, Swam, P_O, P, OAS, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSens, Lag, Lock, PPFq, NxtPP, NxtRst1, PrepWac, PrepCoup, FA2, Const2, Mult2, Rate2, RF2, Floor2, Cap2, PF2, PC2, AB2, NxtRst2, LookBack2, LookBackRate1, LookBackRate2, WamDate, PPUnits, PPCRShift, PayCurrEscr, RcvCurrEscr, BookPrice, AmortMethod, ClientName, BurnStrength )
SELECT Portfolio, PassThruRate, CUSIP, AccountClass, AccountType, ScheduleType, H1, H2, AggNotional, Mult, AggFactor, [Add Accrued?], PV, Swap, AggWac, AggCoup, AggWam, AggAge, Swam, [P/O], AggPrice, OAS, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSense, Lag, Lock, PPFq, NxtPP, NxtRst1, AggPrepWac, AggPrepCoup, FA, Const2, Mult2, Rate, RF, Floor, Cap, PF2, PC, AB, NxtRst2, LookBack, LookBackRate, LookBackRate AS LookBackRate2, WamDate, PPUnits, PPCRShift, PayCurrEscr, RcvCurrEscr, AggBookPrice, AmortMethod, ClientName, 1 AS BurnStrength
FROM tblSecuritiesPalmsSource;

---- qryAppendMPFMarkToModelToPortfolioContentsTableMiddleware
INSERT INTO Portfolio_Contents_MBS_MPF ( Portfolio, IntexID, CUSIP, Account_Class, Account_Type, Sub_Account_Type, H1, H2, Notional, Mult, Factor, Add_Accrued_q, PV, Swap, Wac, Coup, Wam, Age, Swam, P_O, P, OAS, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSens, Lag, Lock, PPFq, NxtPP, NxtRst1, PrepWac, PrepCoup, FA2, Const2, Mult2, Rate2, RF2, Floor2, Cap2, PF2, PC2, AB2, NxtRst2, LookBack2, LookBackRate1, LookBackRate2, WamDate, PPUnits, PPCRShift, PayCurrEscr, RcvCurrEscr, BookPrice, AmortMethod, ClientName, BurnStrength )
SELECT tblMarkToModelLoans.Portfolio, tblMarkToModelLoans.PassThruRate, tblMarkToModelLoans.CUSIP, tblMarkToModelLoans.AccountClass, tblMarkToModelLoans.AccountType, tblMarkToModelLoans.ScheduleType, tblMarkToModelLoans.H1, tblMarkToModelLoans.H2, tblMarkToModelLoans.AggNotional, tblMarkToModelLoans.Mult, tblMarkToModelLoans.AggFactor, tblMarkToModelLoans.[Add Accrued?], tblMarkToModelLoans.PV, tblMarkToModelLoans.Swap, tblMarkToModelLoans.AggWac, tblMarkToModelLoans.AggCoup, tblMarkToModelLoans.AggWam, tblMarkToModelLoans.AggAge, tblMarkToModelLoans.Swam, tblMarkToModelLoans.[P/O], tblMarkToModelLoans.AggPrice, tblMarkToModelLoans.OAS, tblMarkToModelLoans.Settle, tblMarkToModelLoans.Repo, tblMarkToModelLoans.NxtPmt, tblMarkToModelLoans.PP, tblMarkToModelLoans.PPConst, tblMarkToModelLoans.PPMult, tblMarkToModelLoans.PPSense, tblMarkToModelLoans.Lag, tblMarkToModelLoans.Lock, tblMarkToModelLoans.PPFq, tblMarkToModelLoans.NxtPP, tblMarkToModelLoans.NxtRst1, tblMarkToModelLoans.AggPrepWac, tblMarkToModelLoans.AggPrepCoup, tblMarkToModelLoans.FA, tblMarkToModelLoans.Const2, tblMarkToModelLoans.Mult2, tblMarkToModelLoans.Rate, tblMarkToModelLoans.RF, tblMarkToModelLoans.Floor, tblMarkToModelLoans.Cap, tblMarkToModelLoans.PF2, tblMarkToModelLoans.PC, tblMarkToModelLoans.AB, tblMarkToModelLoans.NxtRst2, tblMarkToModelLoans.LookBack, tblMarkToModelLoans.LookBackRate, tblMarkToModelLoans.LookBackRate AS LookBackRate2, tblMarkToModelLoans.WamDate, tblMarkToModelLoans.PPUnits, tblMarkToModelLoans.PPCRShift, tblMarkToModelLoans.PayCurrEscr, tblMarkToModelLoans.RcvCurrEscr, tblMarkToModelLoans.AggBookPrice, tblMarkToModelLoans.AmortMethod, tblMarkToModelLoans.ClientName, 1 AS BurnStrength
FROM tblMarkToModelLoans;

--- qryAppendMPFForwardCommitmentsToInstrumentMiddleware
INSERT INTO Instrument_MBS_MPF ( CUSIP, Active_q, Agency, Delay, Wac, Coup, Balloon, Owam, IF, PF, FA1, Const1, Mult1, Rate1, RF1, Floor1, Cap1, PF1, PC1, AB1, LookBack1, Int_Coup_Mult, PP_Coup_Mult, Sum_Floor, Sum_Cap, Servicing_Model, Loss_Model, Sched_Cap_q, BondCUSIP )
SELECT tblForwardCommitmentPalmsSource.CusipHedged, True AS [Active?], tblForwardCommitmentPalmsSource.Agency, tblForwardCommitmentPalmsSource.Delay, CSng([Wac]) AS [Wac-cnv], CSng([Coup]) AS [Coup-cnv], 0 AS Balloon, CInt([Owam]) AS [Owam-cnv], 1 AS IF, 1 AS PF, tblForwardCommitmentPalmsSource.FA1, CSng([Const1]) AS [Const1-cnv], 1 AS Mult1, tblForwardCommitmentPalmsSource.Rate1, CSng([RF1]) AS [RF1-cnv], -1000000000 AS [Floor1-cnv], 1000000000 AS [Cap1-cnv], 1000000000 AS [PF1-cnv], 1000000000 AS [PC1-cnv], tblForwardCommitmentPalmsSource.AB1, 0 AS LookBack, 0 AS [Int Coup Mult], 1 AS [PP Coup Mult], -10000000000 AS [Sum Floor], 10000000000 AS [Sum Cap], "None" AS [Servicing Model], "None" AS [Loss Model], 0 AS [Sched Cap?], tblForwardCommitmentPalmsSource.Portfolio
FROM tblForwardCommitmentPalmsSource;

--- qryAppendMPFForwardMarkToModelToInstrumentMiddleware
INSERT INTO Instrument_MBS_MPF ( CUSIP, Active_q, Agency, Delay, Wac, Coup, Balloon, Owam, IF, PF, FA1, Const1, Mult1, Rate1, RF1, Floor1, Cap1, PF1, PC1, AB1, LookBack1, Int_Coup_Mult, PP_Coup_Mult, Sum_Floor, Sum_Cap, Servicing_Model, Loss_Model, Sched_Cap_q, BondCUSIP )
SELECT tblForwardCommitmentMarkToModel.CusipHedged, True AS [Active?], tblForwardCommitmentMarkToModel.Agency, 18 AS Delay, CSng([Wac]) AS [Wac-cnv], CSng([Coup]) AS [Coup-cnv], 0 AS Balloon, CInt([Owam]) AS [Owam-cnv], 1 AS IF, 1 AS PF, tblForwardCommitmentMarkToModel.FA1, CSng([Const1]) AS [Const1-cnv], 1 AS Mult1, tblForwardCommitmentMarkToModel.Rate1, CSng([RF1]) AS [RF1-cnv], -1000000000 AS [Floor1-cnv], 1000000000 AS [Cap1-cnv], 1000000000 AS [PF1-cnv], 1000000000 AS [PC1-cnv], tblForwardCommitmentMarkToModel.AB1, 0 AS LookBack, 0 AS [Int Coup Mult], 1 AS [PP Coup Mult], -10000000000 AS [Sum Floor], 10000000000 AS [Sum Cap], "None" AS [Servicing Model], "None" AS [Loss Model], 0 AS [Sched Cap?], tblForwardCommitmentMarkToModel.Portfolio
FROM tblForwardCommitmentMarkToModel;

--- qryAppendMPFForwardCommitmentsToPortfolioContentsMiddleware
INSERT INTO Portfolio_Contents_MBS_MPF ( Portfolio, CUSIP, Account_Class, Account_Type, Sub_Account_Type, H1, H2, Notional, Mult, Factor, Add_Accrued_q, PV, Swap, Wac, Coup, Wam, Age, Swam, P_O, P, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSens, Lag, Lock, PPCRShift, PPFq, NxtPP, NxtRst1, PrepWac, PrepCoup, FA2, Const2, Mult2, Rate2, RF2, Floor2, Cap2, PF2, PC2, AB2, NxtRst2, ClientName, BookPrice, OAS, WamDate, BurnStrength )
SELECT tblForwardCommitmentPalmsSource.Portfolio, tblForwardCommitmentPalmsSource.CusipHedged, tblForwardCommitmentPalmsSource.[Account Class], tblForwardCommitmentPalmsSource.[Account Type], tblForwardCommitmentPalmsSource.[Sub Account Type], tblForwardCommitmentPalmsSource.H1, tblForwardCommitmentPalmsSource.H2, tblForwardCommitmentPalmsSource.Notional, tblForwardCommitmentPalmsSource.Mult, tblForwardCommitmentPalmsSource.Factor, tblForwardCommitmentPalmsSource.[Add Accrued?], tblForwardCommitmentPalmsSource.PV, tblForwardCommitmentPalmsSource.Swap, tblForwardCommitmentPalmsSource.Wac, tblForwardCommitmentPalmsSource.Coup, tblForwardCommitmentPalmsSource.Wam, tblForwardCommitmentPalmsSource.Age, tblForwardCommitmentPalmsSource.Swam, tblForwardCommitmentPalmsSource.[P/O], tblForwardCommitmentPalmsSource.P, tblForwardCommitmentPalmsSource.Settle, tblForwardCommitmentPalmsSource.Repo, tblForwardCommitmentPalmsSource.NxtPmt, tblForwardCommitmentPalmsSource.PP, tblForwardCommitmentPalmsSource.PPConst, tblForwardCommitmentPalmsSource.PPMult, tblForwardCommitmentPalmsSource.PPSens, tblForwardCommitmentPalmsSource.Lag, tblForwardCommitmentPalmsSource.Lock, tblForwardCommitmentPalmsSource.PPCRShift, tblForwardCommitmentPalmsSource.PPFq, tblForwardCommitmentPalmsSource.NxtPP, tblForwardCommitmentPalmsSource.NxtRst1, tblForwardCommitmentPalmsSource.PrepWac, tblForwardCommitmentPalmsSource.PrepCoup, tblForwardCommitmentPalmsSource.FA1 AS FA2, tblForwardCommitmentPalmsSource.Const1 AS Const2, tblForwardCommitmentPalmsSource.Mult1 AS Mult2, tblForwardCommitmentPalmsSource.Rate1 AS Rate2, tblForwardCommitmentPalmsSource.RF1 AS RF2, tblForwardCommitmentPalmsSource.Floor1 AS Floor2, tblForwardCommitmentPalmsSource.Cap1 AS Cap2, tblForwardCommitmentPalmsSource.PF1 AS PF2, tblForwardCommitmentPalmsSource.PC1 AS PC2, tblForwardCommitmentPalmsSource.AB1 AS AB2, tblForwardCommitmentPalmsSource.NxtRst2, tblForwardCommitmentPalmsSource.ClientName, tblForwardCommitmentPalmsSource.Notional AS Book, tblForwardCommitmentPalmsSource.OAS, tblForwardCommitmentPalmsSource.Settle AS WAMDate, 1 AS BurnStrength
FROM tblForwardCommitmentPalmsSource;

---- qryAppendMPFForwardMarkToModelToPortfolioContentsMiddleware
INSERT INTO Portfolio_Contents_MBS_MPF ( Portfolio, CUSIP, Account_Class, Account_Type, Sub_Account_Type, H1, H2, Notional, Mult, Factor, Add_Accrued_q, PV, Swap, Wac, Coup, Wam, Age, Swam, P_O, P, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSens, Lag, Lock, PPCRShift, PPFq, NxtPP, NxtRst1, PrepWac, PrepCoup, FA2, Const2, Mult2, Rate2, RF2, Floor2, Cap2, PF2, PC2, AB2, NxtRst2, ClientName, BookPrice, OAS, WamDate, BurnStrength )
SELECT tblForwardCommitmentMarkToModel.Portfolio, tblForwardCommitmentMarkToModel.CusipHedged, tblForwardCommitmentMarkToModel.[Account Class], tblForwardCommitmentMarkToModel.[Account Type], tblForwardCommitmentMarkToModel.[Sub Account Type], tblForwardCommitmentMarkToModel.H1, tblForwardCommitmentMarkToModel.H2, tblForwardCommitmentMarkToModel.Notional, tblForwardCommitmentMarkToModel.Mult, tblForwardCommitmentMarkToModel.Factor, tblForwardCommitmentMarkToModel.[Add Accrued?], tblForwardCommitmentMarkToModel.PV, tblForwardCommitmentMarkToModel.Swap, tblForwardCommitmentMarkToModel.Wac, tblForwardCommitmentMarkToModel.Coup, tblForwardCommitmentMarkToModel.Wam, tblForwardCommitmentMarkToModel.Age, tblForwardCommitmentMarkToModel.Swam, tblForwardCommitmentMarkToModel.[P/O], tblForwardCommitmentMarkToModel.P, tblForwardCommitmentMarkToModel.Settle, tblForwardCommitmentMarkToModel.Repo, tblForwardCommitmentMarkToModel.NxtPmt, tblForwardCommitmentMarkToModel.PP, tblForwardCommitmentMarkToModel.PPConst, tblForwardCommitmentMarkToModel.PPMult, tblForwardCommitmentMarkToModel.PPSens, tblForwardCommitmentMarkToModel.Lag, tblForwardCommitmentMarkToModel.Lock, tblForwardCommitmentMarkToModel.PPCRShift, tblForwardCommitmentMarkToModel.PPFq, tblForwardCommitmentMarkToModel.NxtPP, tblForwardCommitmentMarkToModel.NxtRst1, tblForwardCommitmentMarkToModel.PrepWac, tblForwardCommitmentMarkToModel.PrepCoup, tblForwardCommitmentMarkToModel.FA1 AS FA2, tblForwardCommitmentMarkToModel.Const1 AS Const2, tblForwardCommitmentMarkToModel.Mult1 AS Mult2, tblForwardCommitmentMarkToModel.Rate1 AS Rate2, tblForwardCommitmentMarkToModel.RF1 AS RF2, tblForwardCommitmentMarkToModel.Floor1 AS Floor2, tblForwardCommitmentMarkToModel.Cap1 AS Cap2, tblForwardCommitmentMarkToModel.PF1 AS PF2, tblForwardCommitmentMarkToModel.PC1 AS PC2, tblForwardCommitmentMarkToModel.AB1 AS AB2, tblForwardCommitmentMarkToModel.NxtRst2, tblForwardCommitmentMarkToModel.ClientName, tblForwardCommitmentMarkToModel.Notional AS Book, tblForwardCommitmentMarkToModel.OAS, tblForwardCommitmentMarkToModel.Settle AS WAMDate, 1 AS BurnStrength
FROM tblForwardCommitmentMarkToModel;

--- qryUpdateMiddlewareLoadStats
EXEC usp_DataLoadStatistics_INSERT 'Daily MPF Loans'

-- ===== cmdInsertMiddleware_Click  ------------- end




---- =============== 20140903 GNII (HFS) begin

------- qryInsertHFSLoansEOD
-- SELECT 	hfs.LoanNumber,
-- 		hfs.LoanInvestmentStatus,
-- 		hfs.ChicagoParticipation,
-- 		hfs.InterestRate,
-- 		hfs.FundingDate,
-- 		hfs.ProductCode,
-- 		hfs.Action_Code,
-- 		LEFT(hfs.ProductCode,2) AS ScheduleType,
-- 		CINT(((hfs.InterestRate*20000)-50)/100)/2 AS PassThruRate,
-- 		Right(hfs.ProductCode,2) AS AccountType,
-- 		NZ(Year(mpf.LoanRecordCreationDate),Year(Date())) AS OriginationYear,
-- 		(ScheduleType & OriginationYear & AccountType & PassThruRate*100) AS CUSIP,
-- 		Cdbl(0) AS Price
-- INTO tblHFSLoansEOD
-- FROM (dbo_UV_HFSLoans_NOSHFD AS hfs 
-- INNER JOIN (SELECT dbo_UV_HFSLoans_NOSHFD.LoanNumber, Max(dbo_UV_HFSLoans_NOSHFD.FundingDate) AS FundingDate FROM dbo_UV_HFSLoans_NOSHFD GROUP BY dbo_UV_HFSLoans_NOSHFD.LoanNumber)  AS unqhfs 
-- 	ON (hfs.LoanNumber = unqhfs.LoanNumber) AND (hfs.FundingDate = unqhfs.FundingDate)) 
-- LEFT JOIN tblMPFLoansAggregateSource mpf ON (hfs.LoanNumber = mpf.LoanNumber)

-- ====== cmdLoadHFSLoansEOD_Click
------- qryInsertHFSLoansEOD    ---- Hua 20141010 Assume loanNumber is the unique ID in dbo_UV_HFSLoans_NOSHFD
SELECT 	hfs.LoanNumber,
		hfs.LoanInvestmentStatus,
		hfs.ChicagoParticipation,
		hfs.InterestRate,
		hfs.FundingDate,
		hfs.ProductCode,
		hfs.Action_Code,
		LEFT(hfs.ProductCode,2) AS ScheduleType,
		INT((hfs.InterestRate*200)+0.5)/2 AS PassThruRate,
		Right(hfs.ProductCode,2) AS AccountType,
		NZ(Year(mpf.LoanRecordCreationDate),Year(Date())) AS OriginationYear,
		(ScheduleType & OriginationYear & AccountType & PassThruRate*100) AS CUSIP,
		Cdbl(0) AS Price
INTO tblHFSLoansEOD
FROM dbo_UV_HFSLoans_NOSHFD AS hfs LEFT JOIN tblMPFLoansAggregateSource AS mpf ON hfs.LoanNumber = mpf.LoanNumber
WHERE LoanInvestmentStatus not in ("03", "06", "09");


-------- qryAppendV2HFVloans ====== Hua 20150210
INSERT INTO tblHFSLoansEOD
select LoanNumber, LOANLoanInvestmentStatus as LoanInvestmentStatus, ParticipationPercent AS ChicagoParticipation, InterestRate, 
	LastFundingEntryDate as FundingDate, ProductCode, NULL AS Action_Code, 
	IIf(Left(ProductCode,2)="GL","GL",
	IIf(RemittanceTypeID=1,"SS",
	IIf(RemittanceTypeID=2,"AA", "MA"))) AS ScheduleType, 
	CInt(InterestRate*200)/2 AS PassThruRate,
	Right(ProductCode,2) AS AccountType,	
	NZ(Year(LoanRecordCreationDate),Year(Date())) AS OriginationYear,
	(ScheduleType & OriginationYear & AccountType & PassThruRate*100) AS CUSIP,
	Cdbl(0) AS Price
FROM dbo_UV_AllMPFLoans_NOSHFD 
WHERE ParticipationOrgKey = 3
and LOANLoanInvestmentStatus = "12"
and MPFBalance>0


-- SELECT 	q.LoanNumber,
-- 		q.LoanInvestmentStatus,
-- 		q.ChicagoParticipation,
-- 		q.InterestRate,
-- 		q.lastFundingEntryDate as FundingDate,
-- 		q.ProductCode,
-- 		NULL AS Action_Code,
-- 		IIf(Left(q.ProductCode,2)="GL","GL",
-- 		IIf(q.RemittanceTypeID=1,"SS",
-- 		IIf(q.RemittanceTypeID=2,"AA",
-- 		IIf(q.RemittanceTypeID=3,"MA"," ")))) AS ScheduleType, 
-- 		CInt(q.InterestRate*200)/2 AS PassThruRate,
-- 		Right(q.ProductCode,2) AS AccountType,
-- 		NZ(Year(q.LoanRecordCreationDate),Year(Date())) AS OriginationYear,
-- 		(ScheduleType & OriginationYear & AccountType & PassThruRate*100) AS CUSIP,
-- 		Cdbl(0) AS Price
-- FROM qryV2HFVloans AS q 
-- WHERE LOANLoanInvestmentStatus = "12"
	

-------- qryAppendV2fairValueLoan ------ Hua 20150210 
-- INSERT INTO tblHFSLoansEOD
-- SELECT 	mpf.LoanNumber,
-- 		t.LOANLoanInvestmentStatus as LoanInvestmentStatus,
-- 		mpf.ChicagoParticipation,
-- 		mpf.InterestRate,
-- 		mpf.lastFundingEntryDate as FundingDate,
-- 		mpf.ProductCode,
-- 		t.Action_Code,
-- 		IIf(Left(mpf.ProductCode,2)="GL","GL",
-- 		IIf(mpf.RemittanceTypeID=1,"SS",
-- 		IIf(mpf.RemittanceTypeID=2,"AA",
-- 		IIf(mpf.RemittanceTypeID=3,"MA"," ")))) AS ScheduleType, 
-- 		CInt(mpf.InterestRate*200)/2 AS PassThruRate,
-- 		Right(mpf.ProductCode,2) AS AccountType,
-- 		NZ(Year(mpf.LoanRecordCreationDate),Year(Date())) AS OriginationYear,
-- 		(ScheduleType & OriginationYear & AccountType & PassThruRate*100) AS CUSIP,
-- 		Cdbl(0) AS Price
-- FROM tbltempmpf AS T inner JOIN tblMPFLoansAggregateSource AS mpf ON t.LoanNumber = mpf.LoanNumber
-- WHERE LOANLoanInvestmentStatus = "12"



---- qryDeleteClosedHFSLoans
DELETE *
FROM tblHFSLoansEOD
WHERE (((tblHFSLoansEOD.LoanInvestmentStatus)="08") AND ((tblHFSLoansEOD.Action_Code) In ("60","65","70","71","72")));


----- qryCheckExistingHFSLoanCount
SELECT tblMPFLoansAggregateSource.ProductCode, Count(tblMPFLoansAggregateSource.LoanNumber) AS CountOfExitingLoans
FROM tblMPFLoansAggregateSource INNER JOIN dbo_UV_HFSLoans_NOSHFD ON tblMPFLoansAggregateSource.LoanNumber = dbo_UV_HFSLoans_NOSHFD.LoanNumber
GROUP BY tblMPFLoansAggregateSource.ProductCode
HAVING (((tblMPFLoansAggregateSource.ProductCode) In ("GL03","GL05")));

---- 
SELECT qryCheckExistingHFSLoanCount.ProductCode, qryCheckExistingHFSLoanCount.CountOfExitingLoans
FROM qryCheckExistingHFSLoanCount;

--- qryCheckNewHFSLoanCount
SELECT tblHFSLoansEOD.ProductCode, Count(tblHFSLoansEOD.LoanNumber) AS CountOfLoanNumber
FROM tblMPFLoansAggregateSource RIGHT JOIN tblHFSLoansEOD ON tblMPFLoansAggregateSource.LoanNumber = tblHFSLoansEOD.LoanNumber
WHERE (((tblMPFLoansAggregateSource.LoanNumber) Is Null))
GROUP BY tblHFSLoansEOD.ProductCode, tblMPFLoansAggregateSource.ProductCode
HAVING (((tblMPFLoansAggregateSource.ProductCode) In ("GL03","GL05")));

--- qryCheckMissingHFSLoanCount
SELECT tblMPFLoansAggregateSource.ProductCode, Count(*) AS CountOfMissingLoans
FROM tblMPFLoansAggregateSource LEFT JOIN tblHFSLoansEOD ON tblMPFLoansAggregateSource.LoanNumber = tblHFSLoansEOD.LoanNumber
WHERE tblHFSLoansEOD.LoanNumber Is Null
GROUP BY tblMPFLoansAggregateSource.ProductCode
HAVING (((tblMPFLoansAggregateSource.ProductCode) In ("GL03","GL05")));

--- qryUpdateHFSLoanPrice
UPDATE tblHFSLoansEOD INNER JOIN MPFPrice ON tblHFSLoansEOD.CUSIP = MPFPrice.CUSIP SET tblHFSLoansEOD.Price = MPFPrice.Price;

---- =============== 20140903 GNII end



----================= pricing 201409
-- qryShowMPFCUSIPS
SELECT s.CUSIP, MPFPrice.CUSIP, MPFPrice.Price AS Expr1
FROM tblSecuritiesPalmsSource as s LEFT JOIN MPFPrice as p ON s.CUSIP = MPFPrice.CUSIP
WHERE (((MPFPrice.CUSIP) Is Null)) OR ((([MPFPrice].[Price]) Is Null))
ORDER BY MPFPrice.CUSIP;

---=========== cmdUpdateMPFPrice_Click update MPF price (from xls to tbl)
-- qryUpdatePricesIntblSecuritiesPalmsSourceTest 
UPDATE tblSecuritiesPalmsSource INNER JOIN MPFPrice ON CUSIP = MPFPrice.CUSIP SET AggPrice = [Price];

-- qryUpdatePricesIntblSecuritiesPalmsTestSourceMarkToModel
UPDATE tblMarkToModelLoans, MPFPrice SET tblMarkToModelLoans.AggPrice = [Price]
WHERE ((([ScheduleType]+[AccountClass]+Left([AccountType],1)+'0'+CStr(([PassThruRate]*100)))=[MPFPrice].[CUSIP]));

-- ========== Collect MPF forward level data
-- qryUpdateMPFForwardPrices
UPDATE ForwardSettleDCs, qryMPFForwardPrice SET ForwardSettleDCs.Price = [qryMPFForwardPrice]![Price]
WHERE ((([qryMPFForwardPrice].[Date])=CDate([ForwardSettleDCs].[DeliveryDate])) AND ((CStr([NoteRate]))=[qryMPFForwardPrice].[Rate]) 
AND (([qryMPFForwardPrice].[Rate])=CStr([ForwardSettleDCs].[NoteRate])) AND ((ForwardSettleDCs.[Schedule Type])=[ProdType]) 
AND ((ForwardSettleDCs.ProductCode)=[CatType]));

-- qryMPFForwardPrice 
SELECT CatType, ProdType, Year, Date, [MBS Coupon], IIf([MPF Rate] Is Null,Null,CStr([MPF Rate]/100)) AS Rate, Day, Price, ServicingRemittanceType
FROM MPFForwardPrice
WHERE ((CatType) Is Not Null) AND ((ProdType) Is Not Null) AND ((Year) Is Not Null) AND ((Date) Is Not Null) AND (([MBS Coupon]) Is Not Null) 
AND ((IIf([MPF Rate] Is Null,Null,CStr([MPF Rate]/100))) Is Not Null) AND ((ServicingRemittanceType) Is Not Null) AND ((Price) Is Not Null);



-- qryDeleteDataFromTblForwardCommitmentPalmsSource
DELETE tblForwardCommitmentPalmsSource.*
FROM tblForwardCommitmentPalmsSource;

-- qryDeleteDataFromTblForwardCommitmentMarkToModelLoans
DELETE tblForwardCommitmentMarkToModel.*
FROM tblForwardCommitmentMarkToModel;


======================================= 20140723 From MPFFHLBDW to tblMPFLoansAggregateSource
--------- qryselectallloans -- old
-- SELECT City, State, PFINumber, MANumber, DeliveryCommitmentNumber, LoanNumber, LoanRecordCreationDate, LastFundingEntryDate, ClosingDate, MPFBalance, LoanAmount, 
-- 	TransactionCode, InterestRate, ProductCode, NumberOfMonths, 
-- 	IIf(Left([ProductCode],2)="GL",CDbl(([InterestRate])-0.0046),IIf([ProductCode]="FX15",CDbl(([InterestRate])-0.0035),CDbl(([InterestRate])-0.0039))) AS NewCoupon, 
-- 	IIf(Left([ProductCode],2)="GL",CDbl(([InterestRate])-([ServicingFee]+[ExcessServicingFee])),CDbl(([InterestRate])-0.0039)) AS CouponOld, 
-- 	CInt([InterestRate]*800)/800 AS PrepaymentInterestRate, FirstPaymentDate, MaturityDate, PIAmount, ScheduleCode, NonBusDayTrnEffDate, RemittanceTypeID, 
-- 	ExcessServicingFee, ServicingFee, BatchID, ParticipationOrgKey, ParticipationPercent, OrigLTV, LOANLoanInvestmentStatus
-- FROM dbo_UV_AllMPFLoans_NOSHFD AS mpf
-- WHERE (((ParticipationOrgKey)=3));

--------- qryselectallloans -- new  Hua added the last filter 20141002
--SELECT City, State, PFINumber, MANumber, DeliveryCommitmentNumber, LoanNumber, LoanRecordCreationDate, LastFundingEntryDate, ClosingDate, MPFBalance, LoanAmount, 
--	TransactionCode, InterestRate, ProductCode, NumberOfMonths, 
--	IIF(Left([ProductCode],2) = "GL", IIF(LOANLoanInvestmentStatus <> "01" And LOANLoanInvestmentStatus <> "03",([InterestRate]),CDbl(([InterestRate]) - 0.0046)),
--		IIF([ProductCode] = "FX15", CDbl(([InterestRate]) - 0.0035),CDbl(([InterestRate]) - 0.0039))) AS NewCoupon, 
--	IIF(Left([ProductCode],2) = "GL", IIF(LOANLoanInvestmentStatus <> "01" And LOANLoanInvestmentStatus <> "03",[InterestRate],CDbl(([InterestRate]) - ([ServicingFee] + [ExcessServicingFee]))), 
--		CDbl(([InterestRate]) - 0.0039)) AS CouponOld, 
--	CInt([InterestRate] * 800) / 800 AS PrepaymentInterestRate, FirstPaymentDate, MaturityDate, PIAmount, ScheduleCode, NonBusDayTrnEffDate, RemittanceTypeID, ExcessServicingFee, 
--	ServicingFee, BatchID, ParticipationOrgKey, ParticipationPercent AS ChicagoParticipation, OrigLTV, LOANLoanInvestmentStatus
--FROM dbo_UV_AllMPFLoans_NOSHFD AS mpf
--WHERE ParticipationOrgKey = 3
--and LOANLoanInvestmentStatus not in ("03", "06", "09");

--------- qryselectallloans -- Hua 20141201 changed newCoupon, couponOld
-- SELECT City, State, PFINumber, MANumber, DeliveryCommitmentNumber, LoanNumber, LoanRecordCreationDate, LastFundingEntryDate, ClosingDate, MPFBalance, LoanAmount, 
-- 	TransactionCode, InterestRate, ProductCode, NumberOfMonths, 
-- 	IIF(Left([ProductCode],2) = "GL", IIF(LOANLoanInvestmentStatus <> "01",([InterestRate]),CDbl(([InterestRate]) - 0.0046)),
-- 		IIF([ProductCode] = "FX15", CDbl(([InterestRate]) - 0.0035),CDbl(([InterestRate]) - 0.0039))) AS NewCoupon, 
-- 	IIF(Left([ProductCode],2) = "GL", IIF(LOANLoanInvestmentStatus <> "01",[InterestRate],CDbl(([InterestRate]) - ([ServicingFee] + [ExcessServicingFee]))), 
-- 		CDbl(([InterestRate]) - 0.0039)) AS CouponOld, 
-- 	CInt([InterestRate] * 800) / 800 AS PrepaymentInterestRate, FirstPaymentDate, MaturityDate, PIAmount, ScheduleCode, NonBusDayTrnEffDate, RemittanceTypeID, ExcessServicingFee, 
-- 	ServicingFee, BatchID, ParticipationOrgKey, ParticipationPercent AS ChicagoParticipation, OrigLTV, LOANLoanInvestmentStatus
-- FROM dbo_UV_AllMPFLoans_NOSHFD AS mpf
-- WHERE ParticipationOrgKey = 3
-- and LOANLoanInvestmentStatus not in ("03", "06", "09")

--------- qryselectallloans -- 20141210 Hua changed newCoupon, CouponOld and filter
SELECT City, State, PFINumber, MANumber, DeliveryCommitmentNumber, LoanNumber, LoanRecordCreationDate, LastFundingEntryDate, ClosingDate, MPFBalance, LoanAmount, 
	TransactionCode, InterestRate, ProductCode, NumberOfMonths, 
	IIF(Left([ProductCode],2) = "GL", IIF(LOANLoanInvestmentStatus <> "01",(INT((InterestRate*200)-0.5)/200),CDbl(([InterestRate]) - 0.0046)),
		IIF([ProductCode] = "FX15", CDbl(([InterestRate]) - 0.0035),CDbl(([InterestRate]) - 0.0039))) AS NewCoupon, 
	IIF(Left([ProductCode],2) = "GL", IIF(LOANLoanInvestmentStatus <> "01", INT((InterestRate*200)-0.5)/200,CDbl(([InterestRate]) - ([ServicingFee] + [ExcessServicingFee]))), 
		CDbl(([InterestRate]) - 0.0039)) AS CouponOld, 
	CInt([InterestRate] * 800) / 800 AS PrepaymentInterestRate, FirstPaymentDate, MaturityDate, PIAmount, ScheduleCode, NonBusDayTrnEffDate, RemittanceTypeID, ExcessServicingFee, 
	ServicingFee, BatchID, ParticipationOrgKey, ParticipationPercent AS ChicagoParticipation, OrigLTV, LOANLoanInvestmentStatus
FROM dbo_UV_AllMPFLoans_NOSHFD AS mpf
WHERE ParticipationOrgKey = 3
and LOANLoanInvestmentStatus <> "09"

--- qryAppendAllLoans --- AI chicagoParticipation
-- INSERT INTO tbltempmpf ( city, state, pfinumber, manumber, deliverycommitmentnumber, loannumber, loanrecordcreationdate, 
--    lastfundingentrydate, closingdate, mpfbalance, loanamount, transactioncode, interestrate, productcode, numberofmonths, 
--    newcoupon, couponold, prepaymentinterestrate, firstpaymentdate, maturitydate, piamount, schedulecode, nonbusdaytrneffdate, 
--    remittancetypeid, excessservicingfee, servicingfee, batchid, chicagoparticipation, OrigLTV )
-- SELECT city, state, pfinumber, manumber, 
--    deliverycommitmentnumber, loannumber, loanrecordcreationdate, 
--    lastfundingentrydate, closingdate, mpfbalance, 
--    loanamount, transactioncode, interestrate, productcode, 
--    numberofmonths, newcoupon, couponold, prepaymentinterestrate, 
--    firstpaymentdate, maturitydate, piamount, schedulecode, 
--    nonbusdaytrneffdate, remittancetypeid, excessservicingfee, 
--    servicingfee, batchid, ParticipationPercent, OrigLTV
-- FROM qryselectallloans;

--- qryAppendAllLoans --- AI chicagoParticipation and GN II loan
INSERT INTO tbltempmpf ( city, state, pfinumber, manumber, deliverycommitmentnumber, loannumber, loanrecordcreationdate, lastfundingentrydate, closingdate, mpfbalance, loanamount, 
	transactioncode, interestrate, productcode, numberofmonths, newcoupon, couponold, prepaymentinterestrate, firstpaymentdate, maturitydate, piamount, schedulecode, nonbusdaytrneffdate, 
	remittancetypeid, excessservicingfee, servicingfee, batchid, chicagoparticipation, OrigLTV, LOANLoanInvestmentStatus )
SELECT city, state, pfinumber, manumber, deliverycommitmentnumber, loannumber, loanrecordcreationdate, lastfundingentrydate, closingdate, mpfbalance, loanamount, 
	transactioncode, interestrate, productcode, numberofmonths, newcoupon, couponold, prepaymentinterestrate, firstpaymentdate, maturitydate, piamount, schedulecode, nonbusdaytrneffdate, 
	remittancetypeid, excessservicingfee, servicingfee, batchid, chicagoparticipation, OrigLTV, LOANLoanInvestmentStatus
FROM qryselectallloans;

-- qryCreateTempActionCode
SELECT tbltempmpf.LoanNumber, dbo_UV_DM_MPF_Daily_Loan.Action_Code INTO tmpActionCode
FROM tbltempmpf INNER JOIN dbo_UV_DM_MPF_Daily_Loan ON tbltempmpf.LoanNumber = dbo_UV_DM_MPF_Daily_Loan.LoanNumber;

-- qryUpdateLoanActionCode
UPDATE tbltempmpf INNER JOIN tmpActionCode ON tbltempmpf.LoanNumber = tmpActionCode.LoanNumber SET tbltempmpf.Action_Code = [tmpActionCode].[Action_Code];

-- qryDeleteClosedLoans
DELETE tblTempMPF.LOANLoanInvestmentStatus, tblTempMPF.Action_Code
FROM tblTempMPF
WHERE (((tblTempMPF.LOANLoanInvestmentStatus)="08") AND ((tblTempMPF.Action_Code) In ("60","65","70","71","72")));

-------- qryselectallloanssteptwo
SELECT m.city, m.state, m.pfinumber, m.manumber, m.deliverycommitmentnumber, m.loannumber, m.loanrecordcreationdate, m.lastfundingentrydate, m.closingdate, m.mpfbalance, 
	m.loanamount, m.transactioncode, m.interestrate, m.productcode, m.newcoupon AS newcoupon2, m.couponold, m.prepaymentinterestrate, m.firstpaymentdate, m.maturitydate, 
	m.piamount, m.numberofmonths, m.schedulecode, m.nonbusdaytrneffdate, m.remittancetypeid, tblmonthenddate.monthenddate, m.excessservicingfee, m.servicingfee, 
	[cefee] + [ceperformance] AS macefee, ma.ceperformance AS ceperformancefee, m.batchid, m.chicagoparticipation, m.OrigLTV
FROM tblmonthenddate, tbltempmpf as m INNER JOIN tbltempma as ma ON (m.manumber = ma.manumber) AND (m.pfinumber = ma.pfinumber)
WHERE (((m.mpfbalance) > 0)
AND ((m.nonbusdaytrneffdate) <= [tblmonthenddate]![monthenddate]));

--- qryAppendAllLoansStepTwo – new
INSERT INTO tblmpfloansaggregatesource ( loannumber, deliverycommitmentnumber, manumber, pfinumber, loanrecordcreationdate, lastfundingentrydate, originalloanclosingdate, 
	mpfbalance, originalamount, transactioncode, interestrate, coupon, couponold, prepaymentinterestrate, firstpaymentdate, maturitydate, piamount, numberofmonths, schedulecode, 
	productcode, portfolioindicator, remittancetypeid, age, chicagoparticipation, currentloanbalance, entrydate, excessservicingfee, servicingfee, cefee, ceperformancefee, OrigLTV )
SELECT a.loannumber, a.deliverycommitmentnumber, a.manumber, a.pfinumber, a.loanrecordcreationdate, a.lastfundingentrydate, a.closingdate AS originalloanclosingdate, 
	a.mpfbalance, a.loanamount, a.transactioncode, a.interestrate, a.newcoupon2, a.couponold, a.prepaymentinterestrate, a.firstpaymentdate, a.maturitydate, a.piamount, 
	a.numberofmonths, a.schedulecode, a.productcode, Iif([batchid] IS NOT NULL,"BATCH","FLOW") AS portfolioindicator, a.remittancetypeid, 
	([numberofmonths] - Iif(Datediff("m",a.[monthenddate],0,0) > [numberofmonths],[numberofmonths], Datediff("m",a.[monthenddate],[maturitydate],0,0))) AS age, 
	a.chicagoparticipation, Iif([sched end prin bal] IS NOT NULL,[sched end prin bal],[mpfbalance]) AS currentloanbalance, 
	a.monthenddate, a.excessservicingfee, a.servicingfee, a.macefee, a.ceperformancefee, a.OrigLTV
FROM qryselectallloanssteptwo as a LEFT JOIN tblnorwestloansforfives as n ON a.loanNumber = n.loannumber
WHERE (((a.chicagoparticipation) > 0)
AND ((Iif([sched end prin bal] IS NOT NULL, [sched end prin bal], [mpfbalance])) > 0));

-- qryMaketbltempMA – 
SELECT PFINumber, MANumber, CEFee, ProgramCode, IIf([CEPerformanceFee] Is Null,0,[CEPerformanceFee]) AS CEPerformance INTO tblTempMA
FROM tblMasterAgreement;

======================================= 20140723 From MPFDW to tblMPFLoansAggregateSource ========= End


-------- the difference between MPFBalance and currentLoanBalance
SELECT count(a.LoanNumber) as cnt,  sum(s.CurrentLoanBalance*s.ChicagoParticipation)/1000000 as cbal, sum(a.MPFBalance*ParticipationPercent)/1000000 as amta, 
sum(s.MPFBalance*ChicagoParticipation)/1000000 as amts
FROM dbo_UV_AllMPFLoans_NOSHFD as a INNER JOIN tblMPFLoansAggregateSource1 as s ON a.LoanNumber = s.LoanNumber
WHERE (((a.ParticipationOrgKey)=3) and LOANLoanInvestmentStatus="01")

cnt	cbal	amta	amts
150555	7537.678116127	12373.0520800772	12374.0920751473








-- old qryAggregateForwardSettleCommitments
-- INSERT INTO tblForwardCommitmentPalmsSource ( ProductCode, RemittanceTypeID, DeliveryYear, Wac, DeliveryMonth, WASettleDay, ScheduleType, ScheduleType2, Delay, Portfolio, CUSIP, CusipHedged, Owam, [Sub Account Type], [Account Type], [Account Class], H1, H2, Notional, Mult, Factor, [Add Accrued?], PV, Swap, Age, Wam, Swam, [P/O], OAS, P, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSens, Lag, Lock, PPCRShift, PPFq, NxtPP, NxtRst1, PrepWac, Coup, CoupOld, PrepCoup, FA1, Const1, Mult1, Rate1, RF1, Floor1, Cap1, PF1, PC1, AB1, NxtRst2, BOOK, ClientName, AggNotional, WAPrice, NumberofCommitments, Agency )
-- SELECT ProductCode, RemittanceTypeID, DeliveryYear, CStr((CInt([NoteRate]*200)/200)*100) AS Wac, DeliveryMonth, CInt(Sum(Abs([UnfundedAmountP])*[DeliveryDay])/Sum(Abs([UnfundedAmountP]))) AS WASettleDay, [Schedule Type], [Schedule Type], Last(Delay) AS LastOfDelay, "MPFForward" AS Portfolio, [Schedule Type] & Right([DeliveryYear],4) & IIf([DeliveryMonth]<10,"0" & CStr([DeliveryMonth]),CStr([DeliveryMonth])) & Mid([ProductCode],3,2) & CDbl([Wac])*1000 AS CUSIP, [Schedule Type] & Right([DeliveryYear],4) & IIf([DeliveryMonth]<10,"0" & CStr([DeliveryMonth]),CStr([DeliveryMonth])) & Mid([ProductCode],3,2) & CDbl([Wac])*1000 AS CUSIPHedged, Mid([ProductCode],3,2)*12 AS Owam, ProductCode AS [Sub Account Type], CStr((CInt([NoteRate]*200)/200)*100) AS [Account Type], CStr([DeliveryYear]) AS [Account Class], CStr([WAPrice]) AS H1, CStr([WASettleDay]) AS H2, CStr(Sum([UnfundedAmountP])) AS Notional, "1" AS Mult, "1" AS Factor, "1" AS [Add Accrued?], "Mid" AS PV, "Mid" AS Swap, "1" AS Age, [OWAM]-1 AS Wam, IIf([OWAM]=180,160,IIf([OWAM]=360,335,IIf([OWAM]=240,220,""))) AS Swam, "OAS" AS [P/O], 0 AS OAS, Sum([deliveryamount]*[price])/Sum([deliveryamount]) AS P, IIf([DeliveryMonth]<10,"0" & CStr([DeliveryMonth]),CStr([DeliveryMonth])) & "/" & IIf([WASettleDay]<10,"0" & CStr([WASettleDay]),CStr([WASettleDay])) & "/" & CStr([DeliveryYear]) AS Settle, DLookUp("[RepoRate]","[tblPALMSRepo]") AS Repo, NextDateFwd([DeliveryYear],[DeliveryMonth],[WASettleDay],[Schedule Type]) AS NxtPmt, "AFT" AS PP, "0" AS PPConst, Last(PPMult) AS LastOfPPMult, "1" AS PPSens, "0" AS Lag, "0" AS Lock, 0 AS PPCRShift, "12" AS PPFq, [NxtPmt] AS NxtPP, [NxtPmt] AS NxtRst1, [Wac] AS PrepWac, IIf(Left(Last([ProductCode]),2)="GL",CStr(CDbl([Wac])-0.46),IIf([ProductCode]="FX15",CStr(CDbl([Wac])-0.35),CStr(CDbl([Wac])-0.39))) AS Coup, IIf(Left(Last([ProductCode]),2)="GL",CStr(CDbl([Wac])-(Last([ServicingFee])+Last([CEFee])+Last([ExcessServicingFee]))*100),CStr(CDbl([Wac])-0.39)) AS CoupOld, IIf(Left(Last([ProductCode]),2)="GL",CStr(CDbl([Wac])-0.46),IIf([ProductCode]="FX15",CStr(CDbl([Wac])-0.35),CStr(CDbl([Wac])-0.39))) AS PrepCoup, "F" AS FA1, "0" AS Const1, "0" AS Mult1, "3L" AS Rate1, "12" AS RF1, "None" AS Floor1, "None" AS Cap1, "None" AS PF1, "None" AS PC1, "Bond" AS AB1, [NxtRst1] AS NxtRst2, CStr(Sum([UnfundedAmountP])) AS BOOK, "MPFForward" AS ClientName, Sum(UnfundedAmountP) AS AggNotional, 0 AS WAPrice, Count(NoteRate) AS NumberofCommitments, Agency
-- FROM [qryFundedandUnfundedDCsbyPFI Unhedged]
-- GROUP BY RemittanceTypeID, DeliveryYear, DeliveryMonth, [Schedule Type], ProductCode, CStr((CInt([NoteRate]*200)/200)*100), Agency, [Schedule Type]
-- HAVING (((CStr(Sum([UnfundedAmountP])))>0));

-- ============ Aggregate MPF Forwards 201404
-- qryAggregateForwardSettleCommitments
INSERT INTO tblForwardCommitmentPalmsSource ( 
	ProductCode, RemittanceTypeID, DeliveryYear, Wac, DeliveryMonth, WASettleDay, ScheduleType, ScheduleType2, Delay, Portfolio, CUSIP, CusipHedged, Owam, [Sub Account Type], 
	[Account Type], [Account Class], H1, H2, Notional, Mult, Factor, [Add Accrued?], PV, Swap, Age, Wam, Swam, [P/O], OAS, P, Settle, Repo, NxtPmt, PP, PPConst, PPMult, 
	PPSens, Lag, Lock, PPCRShift, PPFq, NxtPP, NxtRst1, PrepWac, Coup, CoupOld, PrepCoup, FA1, Const1, Mult1, Rate1, RF1, Floor1, Cap1, PF1, PC1, AB1, NxtRst2, BOOK, 
	ClientName, AggNotional, WAPrice, NumberofCommitments, Agency )
SELECT ProductCode, RemittanceTypeID, DeliveryYear, CStr((CInt([NoteRate]*200)/200)*100) AS Wac, DeliveryMonth, 
	CInt(Sum(Abs([UnfundedAmountP])*[DeliveryDay])/Sum(Abs([UnfundedAmountP]))) AS WASettleDay, [Schedule Type], [Schedule Type], Last(Delay) AS LastOfDelay, "MPFForward" AS Portfolio, 
	[Schedule Type] & Right([DeliveryYear],4) & IIf([DeliveryMonth]<10,"0" & CStr([DeliveryMonth]),CStr([DeliveryMonth])) & Mid([ProductCode],3,2) & CDbl([Wac])*1000 AS CUSIP, 
	[Schedule Type] & Right([DeliveryYear],4) & IIf([DeliveryMonth]<10,"0" & CStr([DeliveryMonth]),CStr([DeliveryMonth])) & Mid([ProductCode],3,2) & CDbl([Wac])*1000 AS CUSIPHedged, 
	
	IIF(Mid([ProductCode],3,2)="03",30,IIF(Mid([ProductCode],3,2)="05",15,Mid([ProductCode],3,2)))*12 AS Owam, ProductCode AS [Sub Account Type], 
	
	CStr((CInt([NoteRate]*200)/200)*100) AS [Account Type], CStr([DeliveryYear]) AS [Account Class], CStr([WAPrice]) AS H1, CStr([WASettleDay]) AS H2, 
	CStr(Sum([UnfundedAmountP])) AS Notional, "1" AS Mult, "1" AS Factor, "1" AS [Add Accrued?], "Mid" AS PV, "Mid" AS Swap, "1" AS Age, [OWAM]-1 AS Wam, 
	IIf([OWAM]=180,160,IIf([OWAM]=360,335,IIf([OWAM]=240,220,""))) AS Swam, "OAS" AS [P/O], 0 AS OAS, Sum([deliveryamount]*[price])/Sum([deliveryamount]) AS P, 
	IIf([DeliveryMonth]<10,"0" & CStr([DeliveryMonth]),CStr([DeliveryMonth])) & "/" & IIf([WASettleDay]<10,"0" & CStr([WASettleDay]),CStr([WASettleDay])) & "/" & CStr([DeliveryYear]) AS Settle, 
	DLookUp("[RepoRate]","[tblPALMSRepo]") AS Repo, NextDateFwd([DeliveryYear],[DeliveryMonth],[WASettleDay],[Schedule Type]) AS NxtPmt, "AFT" AS PP, "0" AS PPConst, 
	Last(PPMult) AS LastOfPPMult, "1" AS PPSens, "0" AS Lag, "0" AS Lock, 0 AS PPCRShift, "12" AS PPFq, [NxtPmt] AS NxtPP, [NxtPmt] AS NxtRst1, [Wac] AS PrepWac, 
	IIf(Left(Last([ProductCode]),2)="GL",CStr(CDbl([Wac])-0.46),IIf([ProductCode]="FX15",CStr(CDbl([Wac])-0.35),CStr(CDbl([Wac])-0.39))) AS Coup, 
	IIf(Left(Last([ProductCode]),2)="GL",CStr(CDbl([Wac])-(Last([ServicingFee])+Last([CEFee])+Last([ExcessServicingFee]))*100),CStr(CDbl([Wac])-0.39)) AS CoupOld, 
	IIf(Left(Last([ProductCode]),2)="GL",CStr(CDbl([Wac])-0.46),IIf([ProductCode]="FX15",CStr(CDbl([Wac])-0.35),CStr(CDbl([Wac])-0.39))) AS PrepCoup, 
	"F" AS FA1, "0" AS Const1, "0" AS Mult1, "3L" AS Rate1, "12" AS RF1, "None" AS Floor1, "None" AS Cap1, "None" AS PF1, "None" AS PC1, "Bond" AS AB1, [NxtRst1] AS NxtRst2, 
	CStr(Sum([UnfundedAmountP])) AS BOOK, "MPFForward" AS ClientName, Sum(UnfundedAmountP) AS AggNotional, 0 AS WAPrice, Count(NoteRate) AS NumberofCommitments, Agency
FROM [qryFundedandUnfundedDCsbyPFI Unhedged]
GROUP BY RemittanceTypeID, DeliveryYear, DeliveryMonth, [Schedule Type], ProductCode, CStr((CInt([NoteRate]*200)/200)*100), Agency, [Schedule Type]
HAVING (((CStr(Sum([UnfundedAmountP])))>0));


-- qryFundedandUnfundedDCsbyPFI Unhedged
SELECT ForwardSettleDCs.*, [MPF Hedges].DeliveryCommitmentNumber AS DCNHedged, [MPF Hedges].HedgeID
FROM ForwardSettleDCs LEFT JOIN [MPF Hedges] ON ForwardSettleDCs.DeliveryCommitmentNumber = [MPF Hedges].DeliveryCommitmentNumber
WHERE ((([MPF Hedges].HedgeID) Is Null));

-- qryFundedandUnfundedDCsbyPFI Hedged
SELECT ForwardSettleDCs.*, [MPF Hedges].DeliveryCommitmentNumber AS DCNHedged, [MPF Hedges].PurchaseGroup, [MPF Hedges].HedgeID, [MPF Hedges].SwapHedgeID
FROM ForwardSettleDCs LEFT JOIN [MPF Hedges] ON ForwardSettleDCs.DeliveryCommitmentNumber = [MPF Hedges].DeliveryCommitmentNumber
WHERE ((([MPF Hedges].HedgeID) Is Not Null));

-- qryFundedandUnfundedDCsbyPFI --  20141024
-- SELECT ProductCode, 
-- 	IIf(Left([ProductCode],2)="GL",IIf(Right([ProductCode],2)="03" OR Right([ProductCode],2)="05","GNMA2","GNMA"),IIf(Left([ProductCode],2)="FX","FNMA"," ")) AS Agency, 
-- 	NoteRate, Fee, CDbl(0) AS Price, DeliveryStatus, ScheduleType, IIf(Left([ProductCode],2)="GL",1,dbo_UV_DCParticipation_MRA_NOSHFD.RemittanceTypeID) AS RemittanceTypeID, 
-- 	IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=2,"AA",IIf([RemittanceTypeID]=3,"MA"," ")))) AS [Schedule Type], 
-- 	IIf([Schedule Type]="MS",18,IIf([Schedule Type]="GL",18,IIf([Schedule Type]="MA",2,IIf([Schedule Type]="AA",48,IIf([Schedule Type]="SS",18,18))))) AS Delay, 
-- 	EntryDate, EntryTime, DeliveryDate, 1 AS DTF, FullName, "1" AS PPMult, PFINumber, MANumber, DeliveryCommitmentNumber, DeliveryAmount, 
-- 	nz([DeliveryAmount])*nz([Participation]) AS DeliveryAmountP, FundedAmount, nz([FundedAmount])*nz([Participation]) AS FundedAmountP, 
-- 	([DeliveryAmount])-([FundedAmount]) AS UnfundedAmount, [DeliveryAmountP]-[FundedAmountP] AS UnfundedAmountP, LastUpdatedDate, ScheduleCode, Participation, 
-- 	Year([DeliveryDate]) AS DeliveryYear, Month([DeliveryDate]) AS DeliveryMonth, Day([DeliveryDate]) AS DeliveryDay, 
-- 	[Schedule Type] & Right([DeliveryYear],4) & IIf([DeliveryMonth]<10,"0" & CStr([DeliveryMonth]),CStr([DeliveryMonth])) & Mid([ProductCode],3,2) & IIF(
-- 		Right([ProductCode],2) = "03" OR Right([ProductCode],2) = "05", CDBL((([NoteRate]*20000)-50)/100)/200, (CInt([NoteRate]*200)/200))*100*1000 AS CUSIP, 
-- 	IsExtended, ServicingFee, ExcessServicingFee, CEFee, CEPerformanceFee,
-- 	IIf(Left([ProductCode],2)="GL",IIF(Right([ProductCode],2) = "03" OR Right([ProductCode],2) = "05", CDBL([NoteRate]),CDbl([NoteRate])-0.0046),
-- 		IIf([ProductCode]="FX15",CDbl([NoteRate])-0.0035,CDbl([NoteRate])-0.0039)) AS Coup, 
-- 	IIf(Left([ProductCode],2)="GL",IIF(Right([ProductCode],2) = "03" OR Right([ProductCode],2) = "05", CDBL([NoteRate]),
-- 		CDbl([NoteRate])-([ServicingFee]+[ExcessServicingFee]+[CEFee])),CDbl([NoteRate])-0.0039) AS CoupOld 
-- INTO ForwardSettleDCs
-- FROM dbo_UV_DCParticipation_MRA_NOSHFD 
-- WHERE ParticipationOrgKey = 3
-- ORDER BY NoteRate, DeliveryDate

-- ======== Load Chicago DC's
-- qryFundedandUnfundedDCsbyPFI --  20141024 Hua changed CUSIP
SELECT ProductCode, 
	IIf(Left([ProductCode],2)="GL",IIf(Right([ProductCode],2)="03" OR Right([ProductCode],2)="05","GNMA2","GNMA"),IIf(Left([ProductCode],2)="FX","FNMA"," ")) AS Agency, 
	NoteRate, Fee, CDbl(0) AS Price, DeliveryStatus, ScheduleType, IIf(Left([ProductCode],2)="GL",1,dbo_UV_DCParticipation_MRA_NOSHFD.RemittanceTypeID) AS RemittanceTypeID, 
	IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=2,"AA",IIf([RemittanceTypeID]=3,"MA"," ")))) AS [Schedule Type], 
	IIf([Schedule Type]="MS",18,IIf([Schedule Type]="GL",18,IIf([Schedule Type]="MA",2,IIf([Schedule Type]="AA",48,IIf([Schedule Type]="SS",18,18))))) AS Delay, 
	EntryDate, EntryTime, DeliveryDate, 1 AS DTF, FullName, "1" AS PPMult, PFINumber, MANumber, DeliveryCommitmentNumber, DeliveryAmount, 
	nz([DeliveryAmount])*nz([Participation]) AS DeliveryAmountP, FundedAmount, nz([FundedAmount])*nz([Participation]) AS FundedAmountP, 
	([DeliveryAmount])-([FundedAmount]) AS UnfundedAmount, [DeliveryAmountP]-[FundedAmountP] AS UnfundedAmountP, LastUpdatedDate, ScheduleCode, Participation, 
	Year([DeliveryDate]) AS DeliveryYear, Month([DeliveryDate]) AS DeliveryMonth, Day([DeliveryDate]) AS DeliveryDay, 
	[Schedule Type] & Right([DeliveryYear],4) & IIf([DeliveryMonth]<10,"0" & CStr([DeliveryMonth]),CStr([DeliveryMonth])) & Mid([ProductCode],3,2) & IIF(
		Right([ProductCode],2) = "03" OR Right([ProductCode],2) = "05", INT(([NoteRate]*200)+0.5)/2, (CInt([NoteRate]*200)/2))*1000 AS CUSIP, 
	IsExtended, ServicingFee, ExcessServicingFee, CEFee, CEPerformanceFee,
	IIf(Left([ProductCode],2)="GL",IIF(Right([ProductCode],2) = "03" OR Right([ProductCode],2) = "05", INT((NoteRate*200)-0.5)/200, CDbl([NoteRate])-0.0046),
		IIf([ProductCode]="FX15",CDbl([NoteRate])-0.0035,CDbl([NoteRate])-0.0039)) AS Coup, 
	IIf(Left([ProductCode],2)="GL",IIF(Right([ProductCode],2) = "03" OR Right([ProductCode],2) = "05", INT((NoteRate*200)-0.5)/200,
		CDbl([NoteRate])-([ServicingFee]+[ExcessServicingFee]+[CEFee])),CDbl([NoteRate])-0.0039) AS CoupOld 
INTO ForwardSettleDCs
FROM dbo_UV_DCParticipation_MRA_NOSHFD 
WHERE ParticipationOrgKey = 3
ORDER BY NoteRate, DeliveryDate



------- qryDeleteMPFandForwardsfromInstrumenttblProduction --- Command59_Click() Not used
DELETE [Instrument-Palms].CUSIP AS Expr1, [Instrument-Palms].BondCUSIP AS Expr2, [Instrument-Palms].[CUSIP], [Instrument-Palms].[BondCUSIP]
FROM [Instrument-Palms]
WHERE ((([Instrument-Palms].[CUSIP]) Like "AA*" Or ([Instrument-Palms].[CUSIP]) Like "GL*" Or ([Instrument-Palms].[CUSIP]) Like "MS*" Or ([Instrument-Palms].[CUSIP]) Like "SS*" Or ([Instrument-Palms].[CUSIP]) Like "DC*" Or ([Instrument-Palms].[CUSIP]) Like "HC*" Or ([Instrument-Palms].[CUSIP]) Like "MA*") AND (([Instrument-Palms].[BondCUSIP])="MPF" Or ([Instrument-Palms].[BondCUSIP])="MPFForward"));


-- ======== Aggregate MPF Forwards
-- AppendForwardSettleDCs MPFHedges
INSERT INTO [MPF Hedges] ( DeliveryCommitmentNumber, Cusip, HedgeID, SwapHedgeID )
SELECT ForwardSettleDCs.DeliveryCommitmentNumber, "HC" & [ForwardSettleDCs].[DeliveryCommitmentNumber] AS Cusip, "HedgeSA" AS HedgeID, "HedgeSA" AS SwapHedgeID
FROM ForwardSettleDCs;


SELECT ForwardSettleDCs.*, [MPF Hedges].DeliveryCommitmentNumber AS DCNHedged, [MPF Hedges].PurchaseGroup, [MPF Hedges].HedgeID, [MPF Hedges].SwapHedgeID
FROM ForwardSettleDCs LEFT JOIN [MPF Hedges] ON ForwardSettleDCs.DeliveryCommitmentNumber = [MPF Hedges].DeliveryCommitmentNumber
WHERE ((([MPF Hedges].HedgeID) Is Not Null));



-- qryUpdatePricesIntblSecuritiesPalmsSource
UPDATE tblSecuritiesPalmsSource INNER JOIN tblFinalPriceSheet ON tblSecuritiesPalmsSource.CUSIP = tblFinalPriceSheet.Cusip SET tblSecuritiesPalmsSource.AggPrice = [New Price]
WHERE (((tblSecuritiesPalmsSource.H1)="Seasoned"));


------------ MPFForwardPrice and MPFPrice are spread sheets
-- qryMPFForwardPrice
SELECT CatType, ProdType, Year, Date, [MBS Coupon], IIf([MPF Rate] Is Null,Null,CStr([MPF Rate]/100)) AS Rate, Day, Price
FROM MPFForwardPrice
WHERE (((CatType) Is Not Null) AND ((ProdType) Is Not Null) AND ((Year) Is Not Null) AND ((Date) Is Not Null) AND (([MBS Coupon]) Is Not Null) AND ((IIf([MPF Rate] Is Null,Null,CStr([MPF Rate]/100))) Is Not Null) AND ((Price) Is Not Null));



------======= Run Palm data report
------ qryPalmsDataSource:
-- SELECT "MPF" AS Portfolio, 
--    IIf(Left([ProductCode],2)="MS","MS",IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=3,"MA","AA")))) AS ScheduleType, 
--    True AS [Active?], 
--    IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]),Year([LoanRecordCreationDate])) AS OriginationYear, 
--    LoanRecordCreationDate, LoanNumber, DeliveryCommitmentNumber, 
--    IIf([PortfolioIndicator]="BATCH",IIf(Month([OriginalLoanClosingDate])>9,Month([OriginalLoanClosingDate]),"0" & Month([OriginalLoanClosingDate])),IIf(Month([LoanRecordCreationDate])>9,Month([LoanRecordCreationDate]),"0" & Month([LoanRecordCreationDate]))) AS OriginationMonth, 
--    Right([ProductCode],2) AS AccountType, 
--    (CInt([InterestRate]*200)/200)*100 AS PassThruRate, 
--    [ScheduleType] & [OriginationYear] & [AccountType] & ([PassThruRate]*100) AS CUSIP, 
--    [OriginationYear] AS AccountClass, [MPFBalance]*[ChicagoParticipation] AS Notional, 
--    1 AS Mult, 
--    IIf([CurrentLoanBalance]=0,[OriginalAmount]/[OriginalAmount],[CurrentLoanBalance]/[OriginalAmount]) AS Factor, 
--    "1" AS [Add Accrued?], "Mid" AS PV, "Mid" AS Swap, 
--    (CInt([InterestRate]*100000)/100000)*100 AS Wac, 
--    (CInt([Coupon]*100000)/100000)*100 AS Coup, 
--    [Wac]-[Coup] AS Diff, 
--    CInt(NPer(([InterestRate]/12),([PIAmount]*-1),[CurrentLoanBalance])) AS WAM, 
--    Age, 
--    IIf([AccountType]=15,160,IIf([AccountType]=20,220,IIf([AccountType]=30,335,""))) AS Swam, 
--    tblPALMSRepo.RepoRate AS Repo, 
--    EntryDate AS Settle, 
--    NextDate([Settle],[ScheduleType]) AS NxtPmt, "0" AS PPConst, "1" AS PPMult, "1" AS PPSense, "0" AS Lag, "0" AS Lock, "12" AS PPFq, [NxtPmt] AS NxtPP, 
--    NextDate([Settle],[ScheduleType]) AS NxtRst1, 
--    [Wac] AS PrepWac, [Coup] AS PrepCoup, "Fixed" AS FA, "0" AS Const2, 1 AS Mult1, 0 AS Mult2, "3L" AS Rate, "12" AS RF, -1000000000 AS Floor, 1000000000 AS Cap, 
--    1 AS PF, 1000000000 AS PF1, 1000000000 AS PF2, 1000000000 AS PC, "Bond" AS AB, [NxtRst1] AS NxtRst2, 0 AS LookBack2, 0 AS LookBackRate, 
--    EntryDate AS WamDate, "CPR" AS PPUnits, 0 AS PPCRShift, 0 AS RcvCurrEscr, 0 AS PayCurrEscr, "StraightLine" AS AmortMethod, "MPFProgram" AS ClientName, 
--    PrepaymentInterestRate, 
--    CurrentLoanBalance AS AcctBalance, 
--    IIf([ScheduleType]="GL","GNMA",IIf([ScheduleType]="MS","GNMA","FNMA")) AS Agency, 
--    IIf([ScheduleType]="MS",18,IIf([ScheduleType]="GL",18,IIf([ScheduleType]="MA",2,IIf([ScheduleType]="AA",48,IIf([ScheduleType]="SS",18,18))))) AS Delay, 
--    NumberOfMonths AS OWAM, 0 AS Ballon, 1 AS IF, 0 AS Const1, 0 AS [Int Coup Mult], 1 AS [PP Coup Mult], -10000000000 AS [Sum Floor], 
--    10000000000 AS [Sum Cap], "None" AS [Servicing Model], "None" AS [Loss Model], 0 AS [Sched Cap?], 
--    CurrentLoanBalance, OriginalAmount, ChicagoParticipation, LoanRecordCreationDate
-- FROM tblMPFLoansAggregateSource, tblPALMSRepo
-- ORDER BY LoanNumber;



	CInt(NPer(([InterestRate]/12),([PIAmount]*-1),[CurrentLoanBalance])) AS WAM, Age, 



SELECT "MPF" AS Portfolio, 
	IIf(Left([ProductCode],2)="MS","MS",IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=3,"MA","AA")))) AS ScheduleType, LoanRecordCreationDate,
	True AS [Active?], 
	IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]), Year([LoanRecordCreationDate])) AS OriginationYear, 
	LoanRecordCreationDate, LoanNumber, DeliveryCommitmentNumber, 
	IIf([PortfolioIndicator]="BATCH", IIf(Month([OriginalLoanClosingDate])>9,Month([OriginalLoanClosingDate]),"0" & Month([OriginalLoanClosingDate])),
		IIf(Month([LoanRecordCreationDate])>9,Month([LoanRecordCreationDate]),"0" & Month([LoanRecordCreationDate]))) AS OriginationMonth, 
	Right([ProductCode],2) AS AccountType, 
	IIF(AccountType = "03" OR AccountType = "05", INT((InterestRate*200)+0.5)/2, (CInt([InterestRate]*200)/2)) AS PassThruRate, 
	[ScheduleType] & [OriginationYear] & [AccountType] & ([PassThruRate]*100) AS CUSIP, [OriginationYear] AS AccountClass, 
	[MPFBalance]*[ChicagoParticipation] AS Notional, 1 AS Mult, 
	IIf([CurrentLoanBalance]=0,[OriginalAmount]/[OriginalAmount],[CurrentLoanBalance]/[OriginalAmount]) AS Factor, 
	"1" AS [Add Accrued?], "Mid" AS PV, "Mid" AS Swap, (CInt([InterestRate]*100000)/100000)*100 AS Wac, (CInt([Coupon]*100000)/100000)*100 AS Coup, 
	[Wac]-[Coup] AS Diff, 
	CInt(IIf(ISERR(NPer(([InterestRate]/12),([PIAmount]*-1),[CurrentLoanBalance])),OriginationMonth,
		NPer(([InterestRate]/12),([PIAmount]*-1),[CurrentLoanBalance]))) AS WAM, 
	Age, 
	IIf([AccountType]=15,160,IIf([AccountType]=20,220,IIf([AccountType]=30,335,IIf([AccountType]=03,335,IIf([AccountType]=05,160,""))))) AS Swam, tblPALMSRepo.RepoRate AS Repo, EntryDate AS Settle, NextDate([Settle],[ScheduleType]) AS NxtPmt, "0" AS PPConst, "1" AS PPMult, "1" AS PPSense, "0" AS Lag, "0" AS Lock, "12" AS PPFq, [NxtPmt] AS NxtPP, NextDate([Settle],[ScheduleType]) AS NxtRst1, [Wac] AS PrepWac, [Coup] AS PrepCoup, "Fixed" AS FA, "0" AS Const2, 1 AS Mult1, 0 AS Mult2, "3L" AS Rate, "12" AS RF, -1000000000 AS Floor, 1000000000 AS Cap, 1 AS PF, 1000000000 AS PF1, 1000000000 AS PF2, 1000000000 AS PC, "Bond" AS AB, [NxtRst1] AS NxtRst2, 0 AS LookBack2, 0 AS LookBackRate, EntryDate AS WamDate, "CPR" AS PPUnits, 0 AS PPCRShift, 0 AS RcvCurrEscr, 0 AS PayCurrEscr, "StraightLine" AS AmortMethod, "MPFProgram" AS ClientName, PrepaymentInterestRate, CurrentLoanBalance AS AcctBalance, IIf([ScheduleType]="GL", IIf(AccountType = "03" OR AccountType = "05","GNMA2","GNMA"),IIf([ScheduleType]="MS","GNMA","FNMA")) AS Agency, IIf([ScheduleType]="MS",18,IIf([ScheduleType]="GL",18,IIf([ScheduleType]="MA",2,IIf([ScheduleType]="AA",48, 18)))) AS Delay, NumberOfMonths AS OWAM, 0 AS Ballon, 1 AS IF, 0 AS Const1, 0 AS [Int Coup Mult], 1 AS [PP Coup Mult], -10000000000 AS [Sum Floor], 10000000000 AS [Sum Cap], "None" AS [Servicing Model], "None" AS [Loss Model], 0 AS [Sched Cap?], CurrentLoanBalance, OriginalAmount, ChicagoParticipation 
FROM tblMPFLoansAggregateSource, tblPALMSRepo
ORDER BY LoanNumber;
















---------- qryPalmsDataSource:  --- Hua 20141208 CHANGED Delay AND PassThruRate
SELECT "MPF" AS Portfolio, 
	IIf(Left([ProductCode],2)="MS","MS",IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=3,"MA","AA")))) AS ScheduleType, 
	True AS [Active?], 
	IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]), Year([LoanRecordCreationDate])) AS OriginationYear, 
	LoanRecordCreationDate, LoanNumber, DeliveryCommitmentNumber, 
	IIf([PortfolioIndicator]="BATCH", IIf(Month([OriginalLoanClosingDate])>9,Month([OriginalLoanClosingDate]),"0" & Month([OriginalLoanClosingDate])),
		IIf(Month([LoanRecordCreationDate])>9,Month([LoanRecordCreationDate]),"0" & Month([LoanRecordCreationDate]))) AS OriginationMonth, 
	Right([ProductCode],2) AS AccountType, 
	IIF(AccountType = "03" OR AccountType = "05", INT((InterestRate*200)+0.5)/2, (CInt([InterestRate]*200)/2)) AS PassThruRate, 
	[ScheduleType] & [OriginationYear] & [AccountType] & ([PassThruRate]*100) AS CUSIP, 
	[OriginationYear] AS AccountClass, [MPFBalance]*[ChicagoParticipation] AS Notional, 
	1 AS Mult, IIf([CurrentLoanBalance]=0,[OriginalAmount]/[OriginalAmount],[CurrentLoanBalance]/[OriginalAmount]) AS Factor, 
	"1" AS [Add Accrued?], "Mid" AS PV, "Mid" AS Swap, (CInt([InterestRate]*100000)/100000)*100 AS Wac, (CInt([Coupon]*100000)/100000)*100 AS Coup, [Wac]-[Coup] AS Diff,
	CInt(NPer(([InterestRate]/12),([PIAmount]*-1),[CurrentLoanBalance])) AS WAM, Age,
	IIf([AccountType]=15,160,IIf([AccountType]=20,220,IIf([AccountType]=30,335,IIf([AccountType]=03,335,IIf([AccountType]=05,160,""))))) AS Swam, 
	tblPALMSRepo.RepoRate AS Repo, 
	EntryDate AS Settle, NextDate([Settle],[ScheduleType]) AS NxtPmt, "0" AS PPConst, "1" AS PPMult, "1" AS PPSense, "0" AS Lag, "0" AS Lock, "12" AS PPFq, [NxtPmt] AS NxtPP,
	NextDate([Settle],[ScheduleType]) AS NxtRst1, [Wac] AS PrepWac, [Coup] AS PrepCoup, "Fixed" AS FA, "0" AS Const2, 1 AS Mult1, 0 AS Mult2, "3L" AS Rate, "12" AS RF, 
	-1000000000 AS Floor, 1000000000 AS Cap, 1 AS PF, 1000000000 AS PF1, 1000000000 AS PF2, 1000000000 AS PC, "Bond" AS AB, [NxtRst1] AS NxtRst2, 0 AS LookBack2, 
	0 AS LookBackRate, EntryDate AS WamDate, "CPR" AS PPUnits, 0 AS PPCRShift, 0 AS RcvCurrEscr, 0 AS PayCurrEscr, "StraightLine" AS AmortMethod, 
	"MPFProgram" AS ClientName, PrepaymentInterestRate, CurrentLoanBalance AS AcctBalance, 
	IIf([ScheduleType]="GL", IIf(AccountType = "03" OR AccountType = "05","GNMA2","GNMA"),IIf([ScheduleType]="MS","GNMA","FNMA")) AS Agency,
	IIf([ScheduleType]="MS",18,IIf([ScheduleType]="GL",18,IIf([ScheduleType]="MA",2,IIf([ScheduleType]="AA",48, 18)))) AS Delay, 
	NumberOfMonths AS OWAM, 0 AS Ballon, 1 AS IF, 0 AS Const1, 0 AS [Int Coup Mult], 1 AS [PP Coup Mult], -10000000000 AS [Sum Floor], 10000000000 AS [Sum Cap], 
	"None" AS [Servicing Model], "None" AS [Loss Model], 0 AS [Sched Cap?], CurrentLoanBalance, OriginalAmount, ChicagoParticipation, LoanRecordCreationDate
FROM tblMPFLoansAggregateSource, tblPALMSRepo
ORDER BY LoanNumber;


-- qryHedgedPalmsDataSource:
SELECT qryPalmsDataSource.*, 
   [MPF Hedges].DeliveryCommitmentNumber AS DeliveryCommitmentNumber, 
   CStr([qryPalmsDataSource].[DeliveryCommitmentNumber]) AS DCNText, 
   "DC" & IIF([SwapHedgeID] is Null, "00", Right([SwapHedgeID], 2)) & [ScheduleType] & Right([AccountClass],4) & [AccountType] & [PassThruRate] AS CusipHedged, 
   [MPF Hedges].PurchaseGroup, 
   [MPF Hedges].HedgeID
FROM qryPalmsDataSource LEFT JOIN [MPF Hedges] ON qryPalmsDataSource.LoanNumber = [MPF Hedges].LoanNumber
WHERE ((([MPF Hedges].HedgeID) Is Not Null));


-- qryUnhedgedPalmsDataSource --
SELECT qryPalmsDataSource.*, [MPF Hedges].DeliveryCommitmentNumber AS DeliveryCommitmentNumber, [MPF Hedges].HedgeID
FROM qryPalmsDataSource LEFT JOIN [MPF Hedges] ON qryPalmsDataSource.LoanNumber = [MPF Hedges].LoanNumber
WHERE ((([MPF Hedges].HedgeID) Is Null));

-- qryAppendUnhedgedDataTotblSecuritiesPalmsSource
INSERT INTO tblSecuritiesPalmsSource ( AccountClass, AccountType, ScheduleType, PassThruRate, Portfolio, CUSIP, H2, AggNotional, Mult, AggFactor, [Add Accrued?], PV, Swap, 
	AggWac, AggCoup, AggWam, AggAge, H1, Swam, AggOWAM, [P/O], AggPrice, OAS, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSense, Lag, Lock, PPFq, NxtPP, NxtRst1, AggPrepWac, 
	AggPrepCoup, FA, Const1, Const2, Rate, RF, Floor, Cap, PF, PF1, PF2, PC, AB, NxtRst2, LookBackRate, LookBack, WamDate, PPUnits, PPCRShift, RcvCurrEscr, PayCurrEscr, 
	AggBookPrice, AmortMethod, ClientName, [Active?], Agency, [Int Coup Mult], [PP Coup Mult], [Sum Floor], [Sum Cap], [Servicing Model], [Loss Model], [Sched Cap?], Delay, 
	Ballon, IF, Mult1, Mult2, PrepaymentInterestRate )
SELECT AccountClass, AccountType, ScheduleType, PassThruRate, 
	Last(Portfolio) AS Portfolio, Last(CUSIP) AS CUSIP, Count(LoanNumber) AS H2, 
	Sum(Notional) AS AggNotional, Last(Mult) AS Mult, IIf(Sum([CurrentLoanBalance])=0,
	Sum([OriginalAmount])/Sum([OriginalAmount]),Sum([CurrentLoanBalance])/Sum([OriginalAmount])) AS AggFactor, Last([Add Accrued?]) AS [Add Accrued?], 
	Last(PV) AS PV, Last(Swap) AS Swap, 
	(Sum([Wac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWac, 
	(Sum([Coup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggCoup, 
	CInt(Sum([Wam]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWam, 
	(Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggAge, 
	IIf(Val([AccountType])=30 And [AggWam]<340,"Seasoned",IIf(Val([AccountType])=20 And [AggWam]<220,"Seasoned",IIf(Val([AccountType])=15 And [AggWam]<160,"Seasoned","MPF"))) AS H1, 
	Last(Swam) AS Swam, CInt(Sum([OWAM]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggOWAM, 
	"OAS" AS [P/O], 0 AS AggPrice, 0 AS OAS, Last(Settle) AS Settle, Last(Repo) AS Repo, 
	Last(NxtPmt) AS NxtPmt, "AFT" AS PP, Last(PPConst) AS PPConst, Last(PPMult) AS PPMult, 
	Last(PPSense) AS PPSense, Last(Lag) AS Lag, Last(Lock) AS Lock, 
	Last(PPFq) AS PPFq, Last(NxtPP) AS NxtPP, Last(NxtRst1) AS NxtRst1, 
	(Sum([PrepWac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepWac, 
	(Sum([PrepCoup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepCoup, 
	Last(FA) AS FA, Last(Const1) AS Const1, Last(Const2) AS Const2, 
	Last(Rate) AS Rate, Last(RF) AS RF, Last(Floor) AS Floor, 
	Last(Cap) AS Cap, Last(PF) AS PF, Last(PF1) AS PF1, 
	Last(PF2) AS PF2, Last(PC) AS PC, Last(AB) AS AB, 
	Last(NxtRst2) AS NxtRst2, Last(LookBackRate) AS LookBackRate1, 
	Last(LookBack2) AS LookBack2, Last(CStr([WamDate])) AS WamDate, 
	Last(PPUnits) AS PPUnits, Avg(PPCRShift) AS PPCRShift, 
	Last(RcvCurrEscr) AS RcvCurrEscr, Last(PayCurrEscr) AS PayCurrEscr, 
	Sum([CurrentLoanBalance]*[ChicagoParticipation]) AS AggBookPrice, Last(AmortMethod) AS AmortMethod, 
	Last(ClientName) AS ClientName, Last([Active?]) AS [Active?], 
	Last(Agency) AS Agency, Last([Int Coup Mult]) AS [Int Coup Mult], 
	Last([PP Coup Mult]) AS [PP Coup Mult], Last([Sum Floor]) AS [Sum Floor], 
	Last([Sum Cap]) AS [Sum Cap], Last([Servicing Model]) AS [Servicing Model], 
	Last([Loss Model]) AS [Loss Model], Last([Sched Cap?]) AS [Sched Cap?], 
	Last(Delay) AS Delay, Last(Ballon) AS Ballon, Last(IF) AS IF, 
	Last(Mult1) AS Mult1, Last(Mult2) AS Mult2, 
	(Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepaymentInterestRate
FROM qryUnhedgedPalmsDataSource
GROUP BY AccountClass, AccountType, ScheduleType, 
	PassThruRate, OriginationYear
HAVING ((((Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])))>30))
ORDER BY OriginationYear, AccountType DESC , ScheduleType DESC , 
	(Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation]));


-- qryAppendUnhedgedDataTotblSecuritiesPalmsSource - NONSEASONED
INSERT INTO tblSecuritiesPalmsSource ( AccountClass, AccountType, ScheduleType, PassThruRate, Portfolio, CUSIP, H2, AggNotional, Mult, AggFactor, [Add Accrued?], PV, Swap, AggWac, AggCoup, 
	AggWam, AggAge, H1, Swam, AggOWAM, [P/O], AggPrice, OAS, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSense, Lag, Lock, PPFq, NxtPP, NxtRst1, AggPrepWac, AggPrepCoup, FA, Const1, 
	Const2, Rate, RF, Floor, Cap, PF, PF1, PF2, PC, AB, NxtRst2, LookBackRate, LookBack, WamDate, PPUnits, PPCRShift, RcvCurrEscr, PayCurrEscr, AggBookPrice, AmortMethod, ClientName, 
	[Active?], Agency, [Int Coup Mult], [PP Coup Mult], [Sum Floor], [Sum Cap], [Servicing Model], [Loss Model], [Sched Cap?], Delay, Ballon, IF, Mult1, Mult2, PrepaymentInterestRate )
SELECT AccountClass, AccountType, ScheduleType, PassThruRate, Last(Portfolio) AS Portfolio, Last(CUSIP) AS CUSIP, Count(LoanNumber) AS H2, Sum(Notional) AS AggNotional, 
	Last(Mult) AS Mult, IIf(Sum([CurrentLoanBalance])=0,Sum([OriginalAmount])/Sum([OriginalAmount]),Sum([CurrentLoanBalance])/Sum([OriginalAmount])) AS AggFactor, 
	Last([Add Accrued?]) AS [Add Accrued?], Last(PV) AS PV, Last(Swap) AS Swap, (Sum([Wac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWac, 
	(Sum([Coup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggCoup, 
	CInt(Sum([Wam]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWam, 
	(Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggAge, 
	IIf(Val([AccountType])=30 And [AggWam]<340,"Seasoned",IIf(Val([AccountType])=20 And [AggWam]<220,"Seasoned",IIf(Val([AccountType])=15 And [AggWam]<160,"Seasoned","MPF"))) AS H1, 
	Last(Swam) AS Swam, CInt(Sum([OWAM]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggOWAM, "OAS" AS [P/O], 0 AS AggPrice, 0 AS OAS, 
	Last(Settle) AS Settle, Last(Repo) AS Repo, Last(NxtPmt) AS NxtPmt, "AFT" AS PP, Last(PPConst) AS PPConst, Last(PPMult) AS PPMult, Last(PPSense) AS PPSense, Last(Lag) AS Lag, 
	Last(Lock) AS Lock, Last(PPFq) AS PPFq, Last(NxtPP) AS NxtPP, Last(NxtRst1) AS NxtRst1, 
	(Sum([PrepWac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepWac, 
	(Sum([PrepCoup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepCoup, 
	Last(FA) AS FA, Last(Const1) AS Const1, Last(Const2) AS Const2, Last(Rate) AS Rate, Last(RF) AS RF, Last(Floor) AS Floor, Last(Cap) AS Cap, Last(PF) AS PF, Last(PF1) AS PF1, 
	Last(PF2) AS PF2, Last(PC) AS PC, Last(AB) AS AB, Last(NxtRst2) AS NxtRst2, Last(LookBackRate) AS LookBackRate1, Last(LookBack2) AS LookBack2, 
	Last(CStr([qryUnhedgedPalmsDataSource].[WamDate])) AS WamDate, Last(PPUnits) AS PPUnits, Avg(PPCRShift) AS PPCRShift, Last(RcvCurrEscr) AS RcvCurrEscr, 
	Last(PayCurrEscr) AS PayCurrEscr, Sum([CurrentLoanBalance]*[ChicagoParticipation]) AS AggBookPrice, Last(AmortMethod) AS AmortMethod, Last(ClientName) AS ClientName, 
	Last([Active?]) AS [Active?], Last(Agency) AS Agency, Last([Int Coup Mult]) AS [Int Coup Mult], Last([PP Coup Mult]) AS [PP Coup Mult], Last([Sum Floor]) AS [Sum Floor], 
	Last([Sum Cap]) AS [Sum Cap], Last([Servicing Model]) AS [Servicing Model], Last([Loss Model]) AS [Loss Model], Last([Sched Cap?]) AS [Sched Cap?], Last(Delay) AS Delay, 
	Last(Ballon) AS Ballon, Last(IF) AS IF, Last(Mult1) AS Mult1, Last(Mult2) AS Mult2, 
	(Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepaymentInterestRate
FROM qryUnhedgedPalmsDataSource
GROUP BY AccountClass, AccountType, ScheduleType, PassThruRate, OriginationYear
HAVING ((((Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])))<=30))
ORDER BY OriginationYear, AccountType DESC , ScheduleType DESC , (Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation]));
	
-- qryAppendHedgedDataTotblSecuritiesPalmsSource
INSERT INTO tblSecuritiesPalmsSource ( AccountClass, AccountType, ScheduleType, PassThruRate, Portfolio, CUSIP, H2, AggNotional, Mult, AggFactor, [Add Accrued?], PV, Swap, AggWac, AggCoup, AggWam, AggAge, H1, Swam, AggOWAM, [P/O], AggPrice, OAS, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSense, Lag, Lock, PPFq, NxtPP, NxtRst1, AggPrepWac, AggPrepCoup, FA, Const1, Const2, Rate, RF, Floor, Cap, PF, PF1, PF2, PC, AB, NxtRst2, LookBackRate, LookBack, WamDate, PPUnits, PPCRShift, RcvCurrEscr, PayCurrEscr, AggBookPrice, AmortMethod, ClientName, [Active?], Agency, [Int Coup Mult], [PP Coup Mult], [Sum Floor], [Sum Cap], [Servicing Model], [Loss Model], [Sched Cap?], Delay, Ballon, IF, Mult1, Mult2, PrepaymentInterestRate )
SELECT qryHedgedPalmsDataSource.AccountClass, qryHedgedPalmsDataSource.AccountType, qryHedgedPalmsDataSource.ScheduleType, qryHedgedPalmsDataSource.PassThruRate, Last(qryHedgedPalmsDataSource.Portfolio) AS Portfolio, Last(qryHedgedPalmsDataSource.CusipHedged) AS CUSIP, Count(qryHedgedPalmsDataSource.LoanNumber) AS H2, Sum(qryHedgedPalmsDataSource.Notional) AS AggNotional, Last(qryHedgedPalmsDataSource.Mult) AS Mult, IIf(Sum([CurrentLoanBalance])=0,Sum([OriginalAmount])/Sum([OriginalAmount]),Sum([CurrentLoanBalance])/Sum([OriginalAmount])) AS AggFactor, Last(qryHedgedPalmsDataSource.[Add Accrued?]) AS [Add Accrued?], Last(qryHedgedPalmsDataSource.PV) AS PV, Last(qryHedgedPalmsDataSource.Swap) AS Swap, (Sum([Wac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWac, (Sum([Coup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggCoup, CInt(Sum([Wam]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWam, (Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggAge, IIf(Val([AccountType])=30 And [AggWam]<340,"Seasoned",IIf(Val([AccountType])=20 And [AggWam]<220,"Seasoned",IIf(Val([AccountType])=15 And [AggWam]<160,"Seasoned","MPF"))) AS H1, Last(qryHedgedPalmsDataSource.Swam) AS Swam, CInt(Sum([OWAM]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggOWAM, "OAS" AS [P/O], 0 AS AggPrice, 0 AS OAS, Last(qryHedgedPalmsDataSource.Settle) AS Settle, Last(qryHedgedPalmsDataSource.Repo) AS Repo, Last(qryHedgedPalmsDataSource.NxtPmt) AS NxtPmt, "AFT" AS PP, Last(qryHedgedPalmsDataSource.PPConst) AS PPConst, Last(qryHedgedPalmsDataSource.PPMult) AS PPMult, Last(qryHedgedPalmsDataSource.PPSense) AS PPSense, Last(qryHedgedPalmsDataSource.Lag) AS Lag, Last(qryHedgedPalmsDataSource.Lock) AS Lock, Last(qryHedgedPalmsDataSource.PPFq) AS PPFq, Last(qryHedgedPalmsDataSource.NxtPP) AS NxtPP, Last(qryHedgedPalmsDataSource.NxtRst1) AS NxtRst1, (Sum([PrepWac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepWac, (Sum([PrepCoup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepCoup, Last(qryHedgedPalmsDataSource.FA) AS FA, Last(qryHedgedPalmsDataSource.Const1) AS Const1, Last(qryHedgedPalmsDataSource.Const2) AS Const2, Last(qryHedgedPalmsDataSource.Rate) AS Rate, Last(qryHedgedPalmsDataSource.RF) AS RF, Last(qryHedgedPalmsDataSource.Floor) AS Floor, Last(qryHedgedPalmsDataSource.Cap) AS Cap, Last(qryHedgedPalmsDataSource.PF) AS PF, Last(qryHedgedPalmsDataSource.PF1) AS PF1, Last(qryHedgedPalmsDataSource.PF2) AS PF2, Last(qryHedgedPalmsDataSource.PC) AS PC, Last(qryHedgedPalmsDataSource.AB) AS AB, Last(qryHedgedPalmsDataSource.NxtRst2) AS NxtRst2, Last(qryHedgedPalmsDataSource.LookBackRate) AS LookBackRate1, Last(qryHedgedPalmsDataSource.LookBack2) AS LookBack2, Last(CStr([qryHedgedPalmsDataSource].[WamDate])) AS WamDate, Last(qryHedgedPalmsDataSource.PPUnits) AS PPUnits, Avg(qryHedgedPalmsDataSource.PPCRShift) AS PPCRShift, Last(qryHedgedPalmsDataSource.RcvCurrEscr) AS RcvCurrEscr, Last(qryHedgedPalmsDataSource.PayCurrEscr) AS PayCurrEscr, Sum([CurrentLoanBalance]*[ChicagoParticipation]) AS AggBookPrice, Last(qryHedgedPalmsDataSource.AmortMethod) AS AmortMethod, Last(qryHedgedPalmsDataSource.ClientName) AS ClientName, Last(qryHedgedPalmsDataSource.[Active?]) AS [Active?], Last(qryHedgedPalmsDataSource.Agency) AS Agency, Last(qryHedgedPalmsDataSource.[Int Coup Mult]) AS [Int Coup Mult], Last(qryHedgedPalmsDataSource.[PP Coup Mult]) AS [PP Coup Mult], Last(qryHedgedPalmsDataSource.[Sum Floor]) AS [Sum Floor], Last(qryHedgedPalmsDataSource.[Sum Cap]) AS [Sum Cap], Last(qryHedgedPalmsDataSource.[Servicing Model]) AS [Servicing Model], Last(qryHedgedPalmsDataSource.[Loss Model]) AS [Loss Model], Last(qryHedgedPalmsDataSource.[Sched Cap?]) AS [Sched Cap?], Last(qryHedgedPalmsDataSource.Delay) AS Delay, Last(qryHedgedPalmsDataSource.Ballon) AS Ballon, Last(qryHedgedPalmsDataSource.IF) AS IF, Last(qryHedgedPalmsDataSource.Mult1) AS Mult1, Last(qryHedgedPalmsDataSource.Mult2) AS Mult2, (Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepaymentInterestRate
FROM qryHedgedPalmsDataSource
GROUP BY qryHedgedPalmsDataSource.AccountClass, qryHedgedPalmsDataSource.AccountType, qryHedgedPalmsDataSource.ScheduleType, qryHedgedPalmsDataSource.PassThruRate, qryHedgedPalmsDataSource.OriginationYear, qryHedgedPalmsDataSource.PurchaseGroup, qryHedgedPalmsDataSource.CusipHedged
HAVING ((((Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])))>30))
ORDER BY qryHedgedPalmsDataSource.OriginationYear, qryHedgedPalmsDataSource.AccountType DESC , qryHedgedPalmsDataSource.ScheduleType DESC , (Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation]));

-- qryAppendHedgedDataTotblSecuritiesPalmsSource – NONSEASONED
INSERT INTO tblSecuritiesPalmsSource ( AccountClass, AccountType, ScheduleType, PassThruRate, Portfolio, CUSIP, H2, AggNotional, Mult, AggFactor, [Add Accrued?], PV, Swap, AggWac, AggCoup, AggWam, AggAge, H1, Swam, AggOWAM, [P/O], AggPrice, OAS, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSense, Lag, Lock, PPFq, NxtPP, NxtRst1, AggPrepWac, AggPrepCoup, FA, Const1, Const2, Rate, RF, Floor, Cap, PF, PF1, PF2, PC, AB, NxtRst2, LookBackRate, LookBack, WamDate, PPUnits, PPCRShift, RcvCurrEscr, PayCurrEscr, AggBookPrice, AmortMethod, ClientName, [Active?], Agency, [Int Coup Mult], [PP Coup Mult], [Sum Floor], [Sum Cap], [Servicing Model], [Loss Model], [Sched Cap?], Delay, Ballon, IF, Mult1, Mult2, PrepaymentInterestRate )
SELECT qryHedgedPalmsDataSource.AccountClass, qryHedgedPalmsDataSource.AccountType, qryHedgedPalmsDataSource.ScheduleType, qryHedgedPalmsDataSource.PassThruRate, Last(qryHedgedPalmsDataSource.Portfolio) AS Portfolio, Last(qryHedgedPalmsDataSource.CusipHedged) AS CUSIP, Count(qryHedgedPalmsDataSource.LoanNumber) AS H2, Sum(qryHedgedPalmsDataSource.Notional) AS AggNotional, Last(qryHedgedPalmsDataSource.Mult) AS Mult, IIf(Sum([CurrentLoanBalance])=0,Sum([OriginalAmount])/Sum([OriginalAmount]),Sum([CurrentLoanBalance])/Sum([OriginalAmount])) AS AggFactor, Last(qryHedgedPalmsDataSource.[Add Accrued?]) AS [Add Accrued?], Last(qryHedgedPalmsDataSource.PV) AS PV, Last(qryHedgedPalmsDataSource.Swap) AS Swap, (Sum([Wac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWac, (Sum([Coup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggCoup, CInt(Sum([Wam]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWam, (Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggAge, IIf(Val([AccountType])=30 And [AggWam]<340,"Seasoned",IIf(Val([AccountType])=20 And [AggWam]<220,"Seasoned",IIf(Val([AccountType])=15 And [AggWam]<160,"Seasoned","MPF"))) AS H1, Last(qryHedgedPalmsDataSource.Swam) AS Swam, CInt(Sum([OWAM]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggOWAM, "OAS" AS [P/O], 0 AS AggPrice, 0 AS OAS, Last(qryHedgedPalmsDataSource.Settle) AS Settle, Last(qryHedgedPalmsDataSource.Repo) AS Repo, Last(qryHedgedPalmsDataSource.NxtPmt) AS NxtPmt, "AFT" AS PP, Last(qryHedgedPalmsDataSource.PPConst) AS PPConst, Last(qryHedgedPalmsDataSource.PPMult) AS PPMult, Last(qryHedgedPalmsDataSource.PPSense) AS PPSense, Last(qryHedgedPalmsDataSource.Lag) AS Lag, Last(qryHedgedPalmsDataSource.Lock) AS Lock, Last(qryHedgedPalmsDataSource.PPFq) AS PPFq, Last(qryHedgedPalmsDataSource.NxtPP) AS NxtPP, Last(qryHedgedPalmsDataSource.NxtRst1) AS NxtRst1, (Sum([PrepWac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepWac, (Sum([PrepCoup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepCoup, Last(qryHedgedPalmsDataSource.FA) AS FA, Last(qryHedgedPalmsDataSource.Const1) AS Const1, Last(qryHedgedPalmsDataSource.Const2) AS Const2, Last(qryHedgedPalmsDataSource.Rate) AS Rate, Last(qryHedgedPalmsDataSource.RF) AS RF, Last(qryHedgedPalmsDataSource.Floor) AS Floor, Last(qryHedgedPalmsDataSource.Cap) AS Cap, Last(qryHedgedPalmsDataSource.PF) AS PF, Last(qryHedgedPalmsDataSource.PF1) AS PF1, Last(qryHedgedPalmsDataSource.PF2) AS PF2, Last(qryHedgedPalmsDataSource.PC) AS PC, Last(qryHedgedPalmsDataSource.AB) AS AB, Last(qryHedgedPalmsDataSource.NxtRst2) AS NxtRst2, Last(qryHedgedPalmsDataSource.LookBackRate) AS LookBackRate1, Last(qryHedgedPalmsDataSource.LookBack2) AS LookBack2, Last(CStr([qryHedgedPalmsDataSource].[WamDate])) AS WamDate, Last(qryHedgedPalmsDataSource.PPUnits) AS PPUnits, Avg(qryHedgedPalmsDataSource.PPCRShift) AS PPCRShift, Last(qryHedgedPalmsDataSource.RcvCurrEscr) AS RcvCurrEscr, Last(qryHedgedPalmsDataSource.PayCurrEscr) AS PayCurrEscr, Sum([CurrentLoanBalance]*[ChicagoParticipation]) AS AggBookPrice, Last(qryHedgedPalmsDataSource.AmortMethod) AS AmortMethod, Last(qryHedgedPalmsDataSource.ClientName) AS ClientName, Last(qryHedgedPalmsDataSource.[Active?]) AS [Active?], Last(qryHedgedPalmsDataSource.Agency) AS Agency, Last(qryHedgedPalmsDataSource.[Int Coup Mult]) AS [Int Coup Mult], Last(qryHedgedPalmsDataSource.[PP Coup Mult]) AS [PP Coup Mult], Last(qryHedgedPalmsDataSource.[Sum Floor]) AS [Sum Floor], Last(qryHedgedPalmsDataSource.[Sum Cap]) AS [Sum Cap], Last(qryHedgedPalmsDataSource.[Servicing Model]) AS [Servicing Model], Last(qryHedgedPalmsDataSource.[Loss Model]) AS [Loss Model], Last(qryHedgedPalmsDataSource.[Sched Cap?]) AS [Sched Cap?], Last(qryHedgedPalmsDataSource.Delay) AS Delay, Last(qryHedgedPalmsDataSource.Ballon) AS Ballon, Last(qryHedgedPalmsDataSource.IF) AS IF, Last(qryHedgedPalmsDataSource.Mult1) AS Mult1, Last(qryHedgedPalmsDataSource.Mult2) AS Mult2, (Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepaymentInterestRate
FROM qryHedgedPalmsDataSource
GROUP BY qryHedgedPalmsDataSource.AccountClass, qryHedgedPalmsDataSource.AccountType, qryHedgedPalmsDataSource.ScheduleType, qryHedgedPalmsDataSource.PassThruRate, qryHedgedPalmsDataSource.OriginationYear, qryHedgedPalmsDataSource.PurchaseGroup, qryHedgedPalmsDataSource.CusipHedged
HAVING ((((Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])))<=30))
ORDER BY qryHedgedPalmsDataSource.OriginationYear, qryHedgedPalmsDataSource.AccountType DESC , qryHedgedPalmsDataSource.ScheduleType DESC , (Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation]));

-- qryHedgedMarkToModel
SELECT qryPalmsDataSource.*, [MPF Hedges].DeliveryCommitmentNumber AS DeliveryCommitmentNumber, CStr([DeliveryCommitmentNumber]) AS DCNText, "HC" & IIF([SwapHedgeID] is Null, "00", Right([SwapHedgeID], 2)) & [ScheduleType] & Right([AccountClass],4) & [AccountType] & [PassThruRate] & "A" AS CusipHedged, [MPF Hedges].PurchaseGroup, [MPF Hedges].OriginalOAS, [MPF Hedges].HedgeID
FROM qryPalmsDataSource LEFT JOIN [MPF Hedges] ON qryPalmsDataSource.LoanNumber = [MPF Hedges].LoanNumber
WHERE ((([MPF Hedges].HedgeID) Is Not Null));


-- qryAppendhedgedDataTotblMarkToModelLoans
INSERT INTO tblMarkToModelLoans ( AccountClass, AccountType, ScheduleType, PassThruRate, Portfolio, CUSIP, H2, AggNotional, Mult, AggFactor, [Add Accrued?], PV, Swap, AggWac, AggCoup, AggWam, AggAge, H1, Swam, AggOWAM, [P/O], AggPrice, OAS, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSense, Lag, Lock, PPFq, NxtPP, NxtRst1, AggPrepWac, AggPrepCoup, FA, Const1, Const2, Rate, RF, Floor, Cap, PF, PF1, PF2, PC, AB, NxtRst2, LookBackRate, LookBack, WamDate, PPUnits, PPCRShift, RcvCurrEscr, PayCurrEscr, AggBookPrice, AmortMethod, ClientName, [Active?], Agency, [Int Coup Mult], [PP Coup Mult], [Sum Floor], [Sum Cap], [Servicing Model], [Loss Model], [Sched Cap?], Delay, Ballon, IF, Mult1, Mult2, PrepaymentInterestRate )
SELECT AccountClass, AccountType, ScheduleType, PassThruRate, "MPFAct" AS Portfolio, Last(CusipHedged) AS CUSIP, Count(LoanNumber) AS H2, Sum(Notional) AS AggNotional, 
	Last(Mult) AS Mult, IIf(Sum([CurrentLoanBalance])=0,Sum([OriginalAmount])/Sum([OriginalAmount]),Sum([CurrentLoanBalance])/Sum([OriginalAmount])) AS AggFactor, 
	Last([Add Accrued?]) AS [Add Accrued?], Last(PV) AS PV, Last(Swap) AS Swap, (Sum([Wac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWac, 
	(Sum([Coup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggCoup, 
	CInt(Sum([Wam]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWam, 
	(Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggAge, 
	IIf(Val([AccountType])=30 And [AggWam]<340,"Seasoned",IIf(Val([AccountType])=20 And [AggWam]<220,"Seasoned",IIf(Val([AccountType])=15 And [AggWam]<160,"Seasoned","MPF"))) AS H1, 
	Last(Swam) AS Swam, CInt(Sum([OWAM]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggOWAM, IIf([H1]="Seasoned","OAS","Price") AS [P/O], 
	0 AS AggPrice, 0 AS OAS, Last(Settle) AS Settle, Last(Repo) AS Repo, Last(NxtPmt) AS NxtPmt, "AFT" AS PP, Last(PPConst) AS PPConst, Last(PPMult) AS PPMult, 
	Last(PPSense) AS PPSense, Last(Lag) AS Lag, Last(Lock) AS Lock, Last(PPFq) AS PPFq, Last(NxtPP) AS NxtPP, Last(NxtRst1) AS NxtRst1, 
	(Sum([PrepWac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepWac, 
	(Sum([PrepCoup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepCoup, Last(FA) AS FA, Last(Const1) AS Const1, 
	Last(Const2) AS Const2, Last(Rate) AS Rate, Last(RF) AS RF, Last(Floor) AS Floor, Last(Cap) AS Cap, Last(PF) AS PF, Last(PF1) AS PF1, Last(PF2) AS PF2, Last(PC) AS PC, 
	Last(AB) AS AB, Last(NxtRst2) AS NxtRst2, Last(LookBackRate) AS LookBackRate1, Last(LookBack2) AS LookBack2, Last(CStr([WamDate])) AS WamDate, 
	Last(PPUnits) AS PPUnits, Avg(PPCRShift) AS PPCRShift, Last(RcvCurrEscr) AS RcvCurrEscr, Last(PayCurrEscr) AS PayCurrEscr, 
	Sum([CurrentLoanBalance]*[ChicagoParticipation]) AS AggBookPrice, Last(AmortMethod) AS AmortMethod, Last(ClientName) AS ClientName, Last([Active?]) AS [Active?], 
	Last(Agency) AS Agency, Last([Int Coup Mult]) AS [Int Coup Mult], Last([PP Coup Mult]) AS [PP Coup Mult], Last([Sum Floor]) AS [Sum Floor], Last([Sum Cap]) AS [Sum Cap], 
	Last([Servicing Model]) AS [Servicing Model], Last([Loss Model]) AS [Loss Model], Last([Sched Cap?]) AS [Sched Cap?], Last(Delay) AS Delay, Last(Ballon) AS Ballon, 
	Last(IF) AS IF, Last(Mult1) AS Mult1, Last(Mult2) AS Mult2, 
	(Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepaymentInterestRate 
FROM qryHedgedMarkToModel
GROUP BY AccountClass, AccountType, ScheduleType, PassThruRate, OriginationYear, PurchaseGroup, CusipHedged
HAVING ((((Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])))>30))
ORDER BY OriginationYear, PurchaseGroup, AccountType DESC , ScheduleType DESC , (Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation]));

-- qryAppendhedgedDataTotblMarkToModelLoans - NONSEASONED
INSERT INTO tblMarkToModelLoans ( AccountClass, AccountType, ScheduleType, PassThruRate, CUSIP, CUSIPOriginal, Portfolio, H2, AggNotional, Mult, AggFactor, [Add Accrued?], PV, Swap, AggWac, AggCoup, AggWam, AggAge, H1, Swam, AggOWAM, [P/O], AggPrice, OAS, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSense, Lag, Lock, PPFq, NxtPP, NxtRst1, AggPrepWac, AggPrepCoup, FA, Const1, Const2, Rate, RF, Floor, Cap, PF, PF1, PF2, PC, AB, NxtRst2, LookBackRate, LookBack, WamDate, PPUnits, PPCRShift, RcvCurrEscr, PayCurrEscr, AggBookPrice, AmortMethod, ClientName, [Active?], Agency, [Int Coup Mult], [PP Coup Mult], [Sum Floor], [Sum Cap], [Servicing Model], [Loss Model], [Sched Cap?], Delay, Ballon, IF, Mult1, Mult2, PrepaymentInterestRate )
SELECT qryHedgedMarkToModel.AccountClass, qryHedgedMarkToModel.AccountType, qryHedgedMarkToModel.ScheduleType, qryHedgedMarkToModel.PassThruRate, qryHedgedMarkToModel.CusipHedged AS CUSIP, qryHedgedMarkToModel.CUSIP AS CUSIPOriginal, "MPFAct" AS Portfolio, Count(qryHedgedMarkToModel.LoanNumber) AS H2, Sum(qryHedgedMarkToModel.Notional) AS AggNotional, Last(qryHedgedMarkToModel.Mult) AS Mult, IIf(Sum([CurrentLoanBalance])=0,Sum([OriginalAmount])/Sum([OriginalAmount]),Sum([CurrentLoanBalance])/Sum([OriginalAmount])) AS AggFactor, Last(qryHedgedMarkToModel.[Add Accrued?]) AS [Add Accrued?], Last(qryHedgedMarkToModel.PV) AS PV, Last(qryHedgedMarkToModel.Swap) AS Swap, (Sum([Wac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWac, (Sum([Coup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggCoup, CInt(Sum([Wam]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWam, (Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggAge, IIf(Val([AccountType])=30 And [AggWam]<340,"Seasoned",IIf(Val([AccountType])=20 And [AggWam]<220,"Seasoned",IIf(Val([AccountType])=15 And [AggWam]<160,"Seasoned","MPF"))) AS H1, Last(qryHedgedMarkToModel.Swam) AS Swam, CInt(Sum([OWAM]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggOWAM, "Price" AS [P/O], 0 AS AggPrice, Avg(qryHedgedMarkToModel.OriginalOAS) AS OAS, Last(qryHedgedMarkToModel.Settle) AS Settle, Last(qryHedgedMarkToModel.Repo) AS Repo, Last(qryHedgedMarkToModel.NxtPmt) AS NxtPmt, "AFT" AS PP, Last(qryHedgedMarkToModel.PPConst) AS PPConst, Last(qryHedgedMarkToModel.PPMult) AS PPMult, Last(qryHedgedMarkToModel.PPSense) AS PPSense, Last(qryHedgedMarkToModel.Lag) AS Lag, Last(qryHedgedMarkToModel.Lock) AS Lock, Last(qryHedgedMarkToModel.PPFq) AS PPFq, Last(qryHedgedMarkToModel.NxtPP) AS NxtPP, Last(qryHedgedMarkToModel.NxtRst1) AS NxtRst1, (Sum([PrepWac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepWac, (Sum([PrepCoup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepCoup, Last(qryHedgedMarkToModel.FA) AS FA, Last(qryHedgedMarkToModel.Const1) AS Const1, Last(qryHedgedMarkToModel.Const2) AS Const2, Last(qryHedgedMarkToModel.Rate) AS Rate, Last(qryHedgedMarkToModel.RF) AS RF, Last(qryHedgedMarkToModel.Floor) AS Floor, Last(qryHedgedMarkToModel.Cap) AS Cap, Last(qryHedgedMarkToModel.PF) AS PF, Last(qryHedgedMarkToModel.PF1) AS PF1, Last(qryHedgedMarkToModel.PF2) AS PF2, Last(qryHedgedMarkToModel.PC) AS PC, Last(qryHedgedMarkToModel.AB) AS AB, Last(qryHedgedMarkToModel.NxtRst2) AS NxtRst2, Last(qryHedgedMarkToModel.LookBackRate) AS LookBackRate1, Last(qryHedgedMarkToModel.LookBack2) AS LookBack2, Last(CStr([qryHedgedMarkToModel].[WamDate])) AS WamDate, Last(qryHedgedMarkToModel.PPUnits) AS PPUnits, Avg(qryHedgedMarkToModel.PPCRShift) AS PPCRShift, Last(qryHedgedMarkToModel.RcvCurrEscr) AS RcvCurrEscr, Last(qryHedgedMarkToModel.PayCurrEscr) AS PayCurrEscr, Sum([CurrentLoanBalance]*[ChicagoParticipation]) AS AggBookPrice, Last(qryHedgedMarkToModel.AmortMethod) AS AmortMethod, Last(qryHedgedMarkToModel.ClientName) AS ClientName, Last(qryHedgedMarkToModel.[Active?]) AS [Active?], Last(qryHedgedMarkToModel.Agency) AS Agency, Last(qryHedgedMarkToModel.[Int Coup Mult]) AS [Int Coup Mult], Last(qryHedgedMarkToModel.[PP Coup Mult]) AS [PP Coup Mult], Last(qryHedgedMarkToModel.[Sum Floor]) AS [Sum Floor], Last(qryHedgedMarkToModel.[Sum Cap]) AS [Sum Cap], Last(qryHedgedMarkToModel.[Servicing Model]) AS [Servicing Model], Last(qryHedgedMarkToModel.[Loss Model]) AS [Loss Model], Last(qryHedgedMarkToModel.[Sched Cap?]) AS [Sched Cap?], Last(qryHedgedMarkToModel.Delay) AS Delay, Last(qryHedgedMarkToModel.Ballon) AS Ballon, Last(qryHedgedMarkToModel.IF) AS IF, Last(qryHedgedMarkToModel.Mult1) AS Mult1, Last(qryHedgedMarkToModel.Mult2) AS Mult2, (Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepaymentInterestRate
FROM qryHedgedMarkToModel
GROUP BY qryHedgedMarkToModel.AccountClass, qryHedgedMarkToModel.AccountType, qryHedgedMarkToModel.ScheduleType, qryHedgedMarkToModel.PassThruRate, qryHedgedMarkToModel.CusipHedged, qryHedgedMarkToModel.CUSIP, qryHedgedMarkToModel.OriginationYear
HAVING ((((Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])))<=30))
ORDER BY qryHedgedMarkToModel.OriginationYear, qryHedgedMarkToModel.AccountType DESC , qryHedgedMarkToModel.ScheduleType DESC , (Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation]));


--------------------------------------------------------- Data for PolyPaths
-- _DlyMPFDC
-- SELECT 'DlyMPFDC' AS Query, "Asset MPFDC " & Left([CUSIP],2) AS Account, 'MBS' AS [Sec Type], 'Asset' AS BSAccount, 'MPFDC' AS SubActI, Left([cusip],2) AS SubActII, 
-- 	tblForwardCommitmentPalmsSource.CUSIP AS [CUST ID], tblForwardCommitmentPalmsSource.CUSIP AS CUSIP, tblForwardCommitmentPalmsSource.Owam AS [Orig Term], 
-- 	tblForwardCommitmentPalmsSource.AggNotional AS Holding, tblForwardCommitmentPalmsSource.Coup AS Coupon, tblForwardCommitmentPalmsSource.Wac AS [WAC(coll)], 
-- 	tblForwardCommitmentPalmsSource.Wam AS [WAM(coll)], tblForwardCommitmentPalmsSource.Age AS [WALA(coll)], 1 AS Factor, tblForwardCommitmentPalmsSource.P AS Price, 
-- 	tblForwardCommitmentPalmsSource.Agency, 'N' AS [Use Static Model], 'USER' AS [Settlement Type], tblForwardCommitmentPalmsSource.Settle AS [Settle Date], '30/360' AS DayCount, 
-- 	IIf(Left([cusip],2)='AA',48,IIf(Left([cusip],2)='MA',2,18)) AS Delay, 'USER' AS Source, 'Price' AS PriceAnchor, 
-- 	'tblForwardCommitmentPalmsSource in N:\Palms\BankDB\MPFFIVES for MiddlewareSQLandNewRemit.mdb' AS PriceSource, 
-- 	IIf(ProductCode="FX30",179,IIf(productCode="FX15",132,IIf(productCode="FX20",143,IIf(productCode="GL30",144,130)))) AS WALoanSize
-- FROM tblForwardCommitmentPalmsSource;


-- ======== Hua 20150106
-- SELECT 'DlyMPFDC' AS Query, "Asset MPFDC " & SubActII AS Account, 'MBS' AS [Sec Type], 'Asset' AS BSAccount, 'MPFDC' AS SubActI, 
-- 	iif(mid(cusip,9,1)="0","GN",Left([cusip],2)) AS SubActII, CUSIP AS [CUST ID], CUSIP, Owam AS [Orig Term], AggNotional AS Holding, 
-- 	Coup AS Coupon, Wac AS [WAC(coll)], Wam AS [WAM(coll)], Age AS [WALA(coll)], 1 AS Factor, P AS Price, Agency, 
-- 	'N' AS [Use Static Model], 'USER' AS [Settlement Type], Settle AS [Settle Date], '30/360' AS DayCount, 
-- 	IIf(Left([cusip],2)='AA',48,IIf(Left([cusip],2)='MA',2,18)) AS Delay, 'USER' AS Source, 'Price' AS PriceAnchor, 
-- 	'tblForwardCommitmentPalmsSource in N:\Palms\BankDB\MPFFIVES for MiddlewareSQLandNewRemit.mdb' AS PriceSource, 
-- 	IIf(ProductCode="FX30",179,IIf(productCode="FX15",132,IIf(productCode="FX20",143,IIf(productCode="GL30",144,130)))) AS WALoanSize
-- FROM tblForwardCommitmentPalmsSource;


-- _DlyMPFDC ========== Hua 20150226 changed subActII and Account
SELECT 'DlyMPFDC' AS Query, "Asset MPFDC " & SubActII AS Account, 'MBS' AS [Sec Type], 'Asset' AS BSAccount, 'MPFDC' AS SubActI, 
	iif(right(ProductCode,2)="03","GN30", IIF(right(ProductCode,2)="05","GN15",ProductCode)) AS SubActII, 
	CUSIP AS [CUST ID], CUSIP, Owam AS [Orig Term], 
	AggNotional AS Holding, Coup AS Coupon, Wac AS [WAC(coll)], 
	Wam AS [WAM(coll)], Age AS [WALA(coll)], 1 AS Factor, P AS Price, 
	Agency, 'N' AS [Use Static Model], 'USER' AS [Settlement Type], Settle AS [Settle Date], '30/360' AS DayCount, 
	IIf(Left([cusip],2)='AA',48,IIf(Left([cusip],2)='MA',2,18)) AS Delay, 'USER' AS Source, 'Price' AS PriceAnchor, 
	'tblForwardCommitmentPalmsSource in N:\Palms\BankDB\MPFFIVES for MiddlewareSQLandNewRemit.mdb' AS PriceSource, 
	IIf(ProductCode="FX30",179,IIf(productCode="FX15",132,IIf(productCode="FX20",143,IIf(productCode="GL30",144,130)))) AS WALoanSize
FROM tblForwardCommitmentPalmsSource;


-- DlyMPF ========== Hua 20150226 changed subActII and Account
SELECT 'DlyMPF' AS Query, "Asset MPF " & SubActII AS Account, 'MBS' AS [Sec Type], 'Asset' AS BSAccount, 'MPF' AS SubActI, 
	Left([CUSIP],2) AS SubActII, 
	tblSecuritiesPalmsSource.CUSIP AS [CUST ID], tblSecuritiesPalmsSource.CUSIP AS CUSIP, tblSecuritiesPalmsSource.AggOWAM AS [Orig Term], 
	[AggBookPrice]-IIf([month 1] Is Null,0,[Month 1]) AS Holding, tblSecuritiesPalmsSource.AggCoup AS Coupon, 
	tblSecuritiesPalmsSource.AggWac AS [WAC(coll)], [AggWam]-IIf([WAM_adj] Is Null,0,[WAM_Adj]) AS [WAM(coll)], 
	tblSecuritiesPalmsSource.AggAge AS [WALA(coll)], tblSecuritiesPalmsSource.Delay, 1 AS Factor, 
	tblSecuritiesPalmsSource.AggPrice AS Price, tblSecuritiesPalmsSource.Agency, 'T+0' AS [Settlement Type], 
	'N' AS [Use Static Model], '30/360' AS DayCount, 'USER' AS Source, 'Price' AS PriceAnchor, 
	'tblSecuritiesPalmsSource in N:\Palms\BankDB\MPFFIVES for MiddlewareSQLandNewRemit.mdb' AS PriceSource, 
	qryWALoanSize_MPF.WALoanSize AS WALoanSize, qryWALoanSize_MPF.LTV
FROM (tblSecuritiesPalmsSource LEFT JOIN MRA_MPFPayDown ON tblSecuritiesPalmsSource.CUSIP=MRA_MPFPayDown.CUSTOM2) 
	LEFT JOIN qryWALoanSize_MPF ON tblSecuritiesPalmsSource.CUSIP=qryWALoanSize_MPF.CUSIP;
	
-- qryWALoanSize_MPF
SELECT qryMPFLoansAggregateSource.CUSIP, Round(Sum(OriginalAmount/1000*cBal)/Sum(cBal),3) AS WALoanSize, Round(Sum(cBal),0) AS CurBal, 
	Round(100*Sum(OrigLTV*cBal)/Sum(cBal),5) AS LTV, qryMPFLoansAggregateSource.AsOfDate
FROM qryMPFLoansAggregateSource
GROUP BY qryMPFLoansAggregateSource.CUSIP, qryMPFLoansAggregateSource.AsOfDate;

-- qryMPFLoansAggregateSource	
SELECT MAS.LoanNumber, IIf(Left([ProductCode],2)="MS","MS",IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=3,"MA","AA")))) AS ScheduleType, 
	IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]),Year([LoanRecordCreationDate])) AS OriginationYear, Right([ProductCode],2) AS AccountType, 
	(CInt([InterestRate]*200)/200)*100 AS PassThruRate, OriginationYear AS AccountClass, 
	IIf(SWAPHedgeID Is Null,[ScheduleType] & [OriginationYear] & [AccountType] & ([PassThruRate]*100),"DC" & IIf(SwapHedgeID Is Null,"00",Right(SwapHedgeID,2)) & ScheduleType & Right(AccountClass,4) & AccountType & PassThruRate) AS CUSIP, 
	OriginalAmount, ChicagoParticipation*CurrentLoanBalance AS CBal, Age, OrigLTV, EntryDate AS AsOfDate
FROM tblMPFLoansAggregateSource AS MAS LEFT JOIN [MPF Hedges] AS H ON H.LoanNumber=MAS.LoanNumber;

-- DlyMPFDC
SELECT Query, Account, [Sec Type], BSAccount, SubActI, SubActII, [CUST ID], CUSIP, [Orig Term], Holding, Coupon, [WAC(coll)], [WAM(coll)], [WALA(coll)], 
	Factor, Price, Agency, [Use Static Model], [Settlement Type], [Settle Date], DayCount, Delay, Source, PriceAnchor, PriceSource, Round([_DlyMPFDC].WALoanSize,3) AS WALoanSize
FROM _DlyMPFDC;

-- FHLBNav_Admin
SELECT FHLBNAV_User.username
FROM FHLBNAV_User
WHERE (((FHLBNAV_User.Admin)=Yes));

-- polypaths Input ======= create the pf file by priceBatch
select [Sec Type],[BSAccount],[SubActI],[CUSIP],'' as [OCUSIP],[SubActII],'' as [Dated Date],'' as [Maturity],[Coupon],[Holding],[DayCount],'' as [Cpn~ Freq~],'' as [SwapCusip],
	'' as [Hedged],'' as [OAS],'' as [Call Date],'' as [Call Price],'' as [Put Date],'' as [Put Price],[Settlement Type],[Source],[Orig Term],[WAC(coll)],[WAM(coll)],[WALA(coll)],
	[Delay],[Factor],[Agency],[Use Static Model],[Price],'' as [Index],'' as [R Mult],'' as [Margin],'' as [First Coupon Date],'' as [Reset Freq~],'' as [DayCount (Pay)],
	'' as [Index(Pay)],'' as [P Mult],'' as [Coupon (Pay)],'' as [First Coupon Date (Pay)],'' as [Cpn~ Freq~ (Pay)],'' as [margin (Pay)],'' as [Reset Freq~ (Pay)],
	'' as [Settle Date],'' as [Strike],'' as [Swaption Swap Effective Date],'' as [Swaption Swap First Pay Date],'' as [Swaption Swap Termination Date],'' as [1st Exer~],
	'' as [Option Type],'' as [Option Exercise Type],'' as [Swaption Strike Rate],'' as [Option Exercise Notice],'' as [jrnlcode],'' as [Opening Market Value],[Cust ID],
	'' as [Sub Type],'' as [DTM],[Account],'' as [Cap],'' as [Floor],[PriceAnchor],[PriceSource],waLoanSize as [Avg Loan Size],[LTV] 
from DlyMPF
UNION 
select [Sec Type],[BSAccount],[SubActI],[CUSIP],'' as [OCUSIP],[SubActII],'' as [Dated Date],'' as [Maturity],[Coupon],[Holding],[DayCount],'' as [Cpn~ Freq~],'' as [SwapCusip],
	'' as [Hedged],'' as [OAS],'' as [Call Date],'' as [Call Price],'' as [Put Date],'' as [Put Price],[Settlement Type],[Source],[Orig Term],[WAC(coll)],[WAM(coll)],[WALA(coll)],
	[Delay],[Factor],[Agency],[Use Static Model],[Price],'' as [Index],'' as [R Mult],'' as [Margin],'' as [First Coupon Date],'' as [Reset Freq~],'' as [DayCount (Pay)],
	'' as [Index(Pay)],'' as [P Mult],'' as [Coupon (Pay)],'' as [First Coupon Date (Pay)],'' as [Cpn~ Freq~ (Pay)],'' as [margin (Pay)],'' as [Reset Freq~ (Pay)],
	[Settle Date],'' as [Strike],'' as [Swaption Swap Effective Date],'' as [Swaption Swap First Pay Date],'' as [Swaption Swap Termination Date],'' as [1st Exer~],
	'' as [Option Type],'' as [Option Exercise Type],'' as [Swaption Strike Rate],'' as [Option Exercise Notice],'' as [jrnlcode],'' as [Opening Market Value],[Cust ID],
	'' as [Sub Type],'' as [DTM],[Account],'' as [Cap],'' as [Floor],[PriceAnchor],[PriceSource],waLoanSize as [Avg Loan Size],'' as [LTV] 
from DlyMPFDC;






-- Update Missing Dates, 1/1/1900, FX15PC,  FX20PC
-- qryUpdateOriginalLoanClosingDateIntblMPFLoanAggregateSource 
UPDATE tblMPFLoansAggregateSource SET tblMPFLoansAggregateSource.OriginalLoanClosingDate = [LastFundingEntryDate] 
WHERE (((tblMPFLoansAggregateSource.OriginalLoanClosingDate)=#1/1/1900#));

-- qryUpdateOriginalLoanClosingDate(MissingData)
UPDATE tblMPFLoansAggregateSource SET tblMPFLoansAggregateSource.OriginalLoanClosingDate = [LastFundingEntryDate] 
WHERE (((tblMPFLoansAggregateSource.OriginalLoanClosingDate) Is Null));

-- qryUpdateProductCodeforFX20bookedthirty
UPDATE tblMPFLoansAggregateSource SET tblMPFLoansAggregateSource.ProductCode = "FX20" 
WHERE (((tblMPFLoansAggregateSource.ProductCode)="FX30") AND ((tblMPFLoansAggregateSource.NumberOfMonths)<=241 And (tblMPFLoansAggregateSource.NumberOfMonths)>=239));


-- ====Run PALMS Data Report
-- qryUpdateFlexiSwap
UPDATE tblFlexiSwap, tblMPFLoansAggregateSource SET tblMPFLoansAggregateSource.ProductCode = "MS" & Right([ProductCode],2);

-- qryUpdateRemitTypeNorthFederalSavingsLoans
UPDATE tblMPFLoansAggregateSource SET tblMPFLoansAggregateSource.RemittanceTypeID = "1"
WHERE (((tblMPFLoansAggregateSource.RemittanceTypeID)=2) AND ((tblMPFLoansAggregateSource.LoanNumber)=134590)) 
OR (((tblMPFLoansAggregateSource.LoanNumber)=134591)) OR (((tblMPFLoansAggregateSource.LoanNumber)=134592)) OR (((tblMPFLoansAggregateSource.LoanNumber)=134593)) 
OR (((tblMPFLoansAggregateSource.LoanNumber)=134594)) OR (((tblMPFLoansAggregateSource.LoanNumber)=134595)) OR (((tblMPFLoansAggregateSource.LoanNumber)=134596)) 
OR (((tblMPFLoansAggregateSource.LoanNumber)=134597)) OR (((tblMPFLoansAggregateSource.LoanNumber)=121576));

-- qryDeleteDataFromtblSecuritiesPalmsSource
DELETE tblSecuritiesPalmsSource.* FROM tblSecuritiesPalmsSource;

-- qryDeleteDataFromtblMarkToModelLoans
DELETE tblMarkToModelLoans.* FROM tblMarkToModelLoans;


---- ========== Export data to Palm Production (not used)









---------- scratch:
'--SELECT "MPF" AS Portfolio, 
'--   IIf(Left([ProductCode],2)="MS","MS",IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=3,"MA","AA")))) AS ScheduleType, 
'--   IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]),Year([LoanRecordCreationDate])) AS OriginationYear, 
'--   LoanNumber,
'--   Right([ProductCode],2) AS AccountType, 
'--   (CInt([InterestRate]*200)/200)*100 AS PassThruRate, 
'--   [ScheduleType] & [OriginationYear] & [AccountType] & ([PassThruRate]*100) AS CUSIP, 
'--   [OriginationYear] AS AccountClass, [MPFBalance]*[ChicagoParticipation] AS Notional, 
'--   IIf([CurrentLoanBalance]=0,[OriginalAmount]/[OriginalAmount],[CurrentLoanBalance]/[OriginalAmount]) AS Factor, 
'--   (CInt([InterestRate]*100000)/100000)*100 AS Wac, 
'--   (CInt([Coupon]*100000)/100000)*100 AS Coup, 
'--   [Wac]-[Coup] AS Diff, 
'--   CInt(NPer(([InterestRate]/12),([PIAmount]*-1),[CurrentLoanBalance])) AS WAM, 
'--   Age, 
'--   IIf([AccountType]=15,160,IIf([AccountType]=20,220,IIf([AccountType]=30,335,""))) AS Swam, 
'--   EntryDate AS Settle, 
'--   PrepaymentInterestRate, 
'--   IIf([ScheduleType]="GL","GNMA",IIf([ScheduleType]="MS","GNMA","FNMA")) AS Agency, 
'--   IIf([ScheduleType]="MS",18,IIf([ScheduleType]="GL",18,IIf([ScheduleType]="MA",2,IIf([ScheduleType]="AA",48,IIf([ScheduleType]="SS",18,18))))) AS Delay, 
'--   CurrentLoanBalance, OriginalAmount, ChicagoParticipation, LoanRecordCreationDate
'--FROM tblMPFLoansAggregateSource
'--ORDER BY LoanNumber;


----- qryLoanCusip (Hua Short)
SELECT 
   IIf(Left([ProductCode],2)="MS","MS",IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=3,"MA","AA")))) AS ScheduleType, 
   IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]),Year([LoanRecordCreationDate])) AS OriginationYear, 
   l.LoanNumber,
   Right([ProductCode],2) AS AccountType, 
   (CInt([InterestRate]*200)/200)*100 AS PassThruRate, 
   IIF (h.LoanNumber IS NULL, [ScheduleType] & [OriginationYear] & [AccountType] & ([PassThruRate]*100), 
      "DC" & IIF([SwapHedgeID] is Null, "00", Right([SwapHedgeID], 2)) & [ScheduleType] & Right([OriginationYear],4) & [AccountType] & [PassThruRate]) AS Cusip,
   [MPFBalance]*[ChicagoParticipation] AS Notional, 
   EntryDate AS Settle, 
   CurrentLoanBalance*ChicagoParticipation AS cBal, 
   OriginalAmount, ChicagoParticipation, PIAmount
FROM tblMPFLoansAggregateSource AS l LEFT JOIN [MPF Hedges] AS h ON l.LoanNumber = h.LoanNumber

-- 20121220 qryLoanCusipNew
SELECT LoanNumber, IIf(Left([ProductCode],2)="MS","MS",IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=3,"MA","AA")))) AS ScheduleType, 
   IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]),Year([LoanRecordCreationDate])) AS OriginationYear, Right([ProductCode],2) AS AccountType, 
   (CInt([InterestRate]*200)/200)*100 AS PassThruRate, [ScheduleType] & [OriginationYear] & [AccountType] & ([PassThruRate]*100) AS Cusip, 
   [MPFBalance]*[ChicagoParticipation] AS Notional, (CInt([InterestRate]*100000)/100000)*100 AS Wac, (CInt([Coupon]*100000)/100000)*100 AS Coup, 
   CInt(NPer(([InterestRate]/12),([PIAmount]*-1),[CurrentLoanBalance])) AS WAM, Age, 
   IIf([AccountType]=15,160,IIf([AccountType]=20,220,IIf([AccountType]=30,335,""))) AS Swam, EntryDate AS Settle, CurrentLoanBalance*ChicagoParticipation AS cBal, 
   OriginalAmount, ChicagoParticipation, CEFee, DeliveryCommitmentNumber, PIAmount, interestRate, maNumber, PFINumber, ProductCode
FROM tblMPFLoansAggregateSource AS l;


---------- qryPalmsDataSource: Old
SELECT "MPF" AS Portfolio, 
   IIf(Left([ProductCode],2)="MS","MS",IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=3,"MA","AA")))) AS ScheduleType, 
   True AS [Active?], 
   IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]),Year([LoanRecordCreationDate])) AS OriginationYear, 
   LoanRecordCreationDate, LoanNumber, DeliveryCommitmentNumber, 
   IIf([PortfolioIndicator]="BATCH",IIf(Month([OriginalLoanClosingDate])>9,Month([OriginalLoanClosingDate]),"0" & Month([OriginalLoanClosingDate])),IIf(Month([LoanRecordCreationDate])>9,Month([LoanRecordCreationDate]),"0" & Month([LoanRecordCreationDate]))) AS OriginationMonth, 
   Right([ProductCode],2) AS AccountType, 
   (CInt([InterestRate]*200)/200)*100 AS PassThruRate, 
   [ScheduleType] & [OriginationYear] & [AccountType] & ([PassThruRate]*100) AS CUSIP, 
   [OriginationYear] AS AccountClass, [MPFBalance]*[ChicagoParticipation] AS Notional, 
   1 AS Mult, 
   IIf([CurrentLoanBalance]=0,[OriginalAmount]/[OriginalAmount],[CurrentLoanBalance]/[OriginalAmount]) AS Factor, 
   "1" AS [Add Accrued?], "Mid" AS PV, "Mid" AS Swap, 
   (CInt([InterestRate]*100000)/100000)*100 AS Wac, 
   (CInt([Coupon]*100000)/100000)*100 AS Coup, 
   [Wac]-[Coup] AS Diff, 
   CInt(NPer(([InterestRate]/12),([PIAmount]*-1),[CurrentLoanBalance])) AS WAM, 
   Age, 
   IIf([AccountType]=15,160,IIf([AccountType]=20,220,IIf([AccountType]=30,335,""))) AS Swam, 
   tblPALMSRepo.RepoRate AS Repo, 
   EntryDate AS Settle, 
   NextDate([Settle],[ScheduleType]) AS NxtPmt, "0" AS PPConst, "1" AS PPMult, "1" AS PPSense, "0" AS Lag, "0" AS Lock, "12" AS PPFq, [NxtPmt] AS NxtPP, 
   NextDate([Settle],[ScheduleType]) AS NxtRst1, 
   [Wac] AS PrepWac, [Coup] AS PrepCoup, "Fixed" AS FA, "0" AS Const2, 1 AS Mult1, 0 AS Mult2, "3L" AS Rate, "12" AS RF, -1000000000 AS Floor, 1000000000 AS Cap, 
   1 AS PF, 1000000000 AS PF1, 1000000000 AS PF2, 1000000000 AS PC, "Bond" AS AB, [NxtRst1] AS NxtRst2, 0 AS LookBack2, 0 AS LookBackRate, 
   EntryDate AS WamDate, "CPR" AS PPUnits, 0 AS PPCRShift, 0 AS RcvCurrEscr, 0 AS PayCurrEscr, "StraightLine" AS AmortMethod, "MPFProgram" AS ClientName, 
   PrepaymentInterestRate, 
   CurrentLoanBalance AS AcctBalance, 
   IIf([ScheduleType]="GL","GNMA",IIf([ScheduleType]="MS","GNMA","FNMA")) AS Agency, 
   IIf([ScheduleType]="MS",18,IIf([ScheduleType]="GL",18,IIf([ScheduleType]="MA",2,IIf([ScheduleType]="AA",48,IIf([ScheduleType]="SS",18,18))))) AS Delay, 
   NumberOfMonths AS OWAM, 0 AS Ballon, 1 AS IF, 0 AS Const1, 0 AS [Int Coup Mult], 1 AS [PP Coup Mult], -10000000000 AS [Sum Floor], 
   10000000000 AS [Sum Cap], "None" AS [Servicing Model], "None" AS [Loss Model], 0 AS [Sched Cap?], 
   CurrentLoanBalance, OriginalAmount, ChicagoParticipation, LoanRecordCreationDate
FROM tblMPFLoansAggregateSource, tblPALMSRepo
ORDER BY LoanNumber;





-- remittance type adjustment
=IF($A74="","",IF(LEFT($D74,2)="AA","S/R",IF(LEFT($D74,2)="MA","A/A","S/S")))
=IF($A74="","",IF($H74="S/S",0,IF($H74="S/R",-0.33,0.02)))

-- qryChicagoParticipation
SELECT tblDeliveryCommitment.DeliveryCommitmentNumber, tblDeliveryCommitment.MANumber, (1-([NewYorkParticipation]+[IndianapolisParticipation]+[BostonParticipation]+[PittsburghParticipation]+[AtlantaParticipation]+[CincinnatiParticipation]+[DesMoinesParticipation]+[TopekaParticipation]+[DallasParticipation]+[SanFranciscoParticipation]+[SeattleParticipation])) AS ChicagoParticipation
FROM tblDeliveryCommitment
GROUP BY tblDeliveryCommitment.DeliveryCommitmentNumber, tblDeliveryCommitment.MANumber, tblDeliveryCommitment.CincinnatiParticipation, tblDeliveryCommitment.NewYorkParticipation, tblDeliveryCommitment.IndianapolisParticipation, tblDeliveryCommitment.BostonParticipation, tblDeliveryCommitment.PittsburghParticipation, tblDeliveryCommitment.AtlantaParticipation, tblDeliveryCommitment.DesMoinesParticipation, tblDeliveryCommitment.TopekaParticipation, tblDeliveryCommitment.DallasParticipation, tblDeliveryCommitment.SanFranciscoParticipation, tblDeliveryCommitment.SeattleParticipation
ORDER BY tblDeliveryCommitment.CincinnatiParticipation;




---------- Old --- qryFlowToAggregateUsingPassthroughSQL
INSERT INTO tblMPFLoansAggregateSource ( 
	LoanNumber, DeliveryCommitmentNumber, MANumber, PFINumber, LoanRecordCreationDate, LastFundingEntryDate, OriginalLoanClosingDate, MPFBalance, OriginalAmount, 
	TransactionCode, InterestRate, Coupon, CouponOld, PrepaymentInterestRate, FirstPaymentDate, MaturityDate, PIAmount, NumberOfMonths, ScheduleCode, ProductCode, 
	PortfolioIndicator, RemittanceTypeID, Age, ChicagoParticipation, CurrentLoanBalance, EntryDate, CEFee, ExcessServicingFee, ServicingFee )
SELECT f.LoanNumber, f.DeliveryCommitmentNumber, f.MANumber, f.PFINumber, f.LoanRecordCreationDate, f.LastFundingEntryDate, f.ClosingDate AS OriginalLoanClosingDate, 
	f.MPFBalance, f.Amount AS Expr1, f.TransactionCode, f.InterestRate, f.NewCoupon, f.CouponOld, f.PrePaymentInterestRate, f.FirstPaymentDate, f.MaturityDate, 
	f.PIAmount, f.NumberOfMonths, f.ScheduleCode, f.ProductCode, "FLOW" AS PortfolioIndicator, f.RemittanceTypeID, 
	([NumberOfMonths]-IIf(DateDiff("m",Date(),[MaturityDate],0,0)>[NumberOfMonths],[NumberOfMonths],DateDiff("m",Date(),[MaturityDate],0,0))) AS Age, 
	f.ChicagoParticipation AS Expr2, IIf([Sched End Prin Bal] Is Not Null,[Sched End Prin Bal],[MPFBalance]) AS CurrentLoanBalance, Date() AS Expr20, 
	f.MACEFEE, f.ExcessServicingFee, f.ServicingFee
FROM tblNorwestLoansForFives as n RIGHT JOIN qryFlowSQL as f ON n.LoanNumber = f.LoanNumber
WHERE (((IIf([Sched End Prin Bal] Is Not Null,[Sched End Prin Bal],[MPFBalance]))<>0))
ORDER BY f.LoanNumber;

----- qryFlowSQL
Select * from UV_MPFFives_Flow_Loans_NOSHFD


--------- qryBatchtoAggregateUsingPassThroughSQL
INSERT INTO tblMPFLoansAggregateSource ( 
	LoanNumber, DeliveryCommitmentNumber, MANumber, PFINumber, LoanRecordCreationDate, LastFundingEntryDate, OriginalLoanClosingDate, MPFBalance, OriginalAmount, 
	TransactionCode, InterestRate, Coupon, CouponOld, PrepaymentInterestRate, FirstPaymentDate, MaturityDate, PIAmount, NumberOfMonths, ScheduleCode, ProductCode, 
	PortfolioIndicator, RemittanceTypeID, Age, ChicagoParticipation, CurrentLoanBalance, EntryDate, CEFee, ExcessServicingFee, ServicingFee )
SELECT DISTINCT b.LoanNumber, b.DeliveryCommitmentNumber, b.MANumber, b.PFINumber, b.LoanRecordCreationDate, b.LastFundingEntryDate, b.ClosingDate AS OriginalLoanClosingDate, 
	b.MPFBalance, b.LoanAmount, b.TransactionCode, b.InterestRate, b.NewCoupon, b.CouponOld, b.PrePaymentInterestRate, b.FirstPaymentDate, b.MaturityDate, b.PIAmount, 
	b.NumberOfMonths, b.ScheduleCode, b.ProductCode, "BATCH" AS PortfolioIndicator, b.RemittanceTypeID, 
	([NumberOfMonths]-IIf(DateDiff("m",Date(),[MaturityDate],0,0)>[NumberOfMonths],[NumberOfMonths],DateDiff("m",Date(),[MaturityDate],0,0))) AS Age, 
	b.ChicagoParticipation, IIf([Sched End Prin Bal] Is Not Null,[Sched End Prin Bal],[MPFBalance]) AS CurrentLoanBalance, Date() AS Expr20, 
	CDbl(b.[MACEFee]) AS MACEFee, b.ExcessServicingFee, b.ServicingFee
FROM tblNorwestLoansForFives as n INNER JOIN qryBatchSQL as b ON n.LoanNumber = b.LoanNumber
WHERE (((IIf([Sched End Prin Bal] Is Not Null,[Sched End Prin Bal],[MPFBalance]))<>0));


--------- qryBatchSQL
Select * from UV_MPFFives_Batch_Loans_NOSHFD



-- qryLoanCusip (Old Hua)
-- SELECT LoanNumber, CusipHedged as Cusip, ScheduleType, OriginationYear, AccountType, PassThruRate, Wac, Coup, WAM, Age, ChicagoParticipation*CurrentLoanBalance as cBal, OriginalAmount, ChicagoParticipation, settle as asOfDate
-- FROM qryHedgedPalmsDataSource
-- UNION SELECT LoanNumber, Cusip, ScheduleType, OriginationYear, AccountType, PassThruRate, Wac, Coup, WAM, Age, ChicagoParticipation*CurrentLoanBalance as cBal, OriginalAmount, ChicagoParticipation, settle as asOfDate
-- FROM qryUnhedgedPalmsDataSource;

qryLoanSize (Hua)
SELECT Cusip, Sum(OriginalAmount/1000*cBal)/Sum(cBal) AS WAloanSize, settle, Sum(cBal) AS curBal
FROM qryLoanCusip
GROUP BY Cusip, settle;

qryFwdLoanSize (Hua)
SELECT CUSIP, IIF(ProductCode="FX30", 179, IIF(productCode="FX15", 132, IIF(productCode="FX20", 143, IIF(productCode="GL30", 144, 130)))) as WALoanSize
FROM tblForwardCommitmentPalmsSource;

qryLoanCusip (Hua)
SELECT "MPF" AS Portfolio, 
   IIf(Left([ProductCode],2)="MS","MS",IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=3,"MA","AA")))) AS ScheduleType, 
   IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]),Year([LoanRecordCreationDate])) AS OriginationYear, 
   l.LoanNumber,
   Right([ProductCode],2) AS AccountType, 
   (CInt([InterestRate]*200)/200)*100 AS PassThruRate, 
   IIF (h.LoanNumber IS NULL, [ScheduleType] & [OriginationYear] & [AccountType] & ([PassThruRate]*100), 
      "DC" & IIF([SwapHedgeID] is Null, "00", Right([SwapHedgeID], 2)) & [ScheduleType] & Right([AccountClass],4) & [AccountType] & [PassThruRate]) AS Cusip,
   [OriginationYear] AS AccountClass, [MPFBalance]*[ChicagoParticipation] AS Notional, 
   IIf([CurrentLoanBalance]=0,[OriginalAmount]/[OriginalAmount],[CurrentLoanBalance]/[OriginalAmount]) AS Factor, 
   (CInt([InterestRate]*100000)/100000)*100 AS Wac, 
   (CInt([Coupon]*100000)/100000)*100 AS Coup, 
   CInt(NPer(([InterestRate]/12),([PIAmount]*-1),[CurrentLoanBalance])) AS WAM, 
   Age, 
   IIf([AccountType]=15,160,IIf([AccountType]=20,220,IIf([AccountType]=30,335,""))) AS Swam, 
   EntryDate AS Settle, 
   IIf([ScheduleType]="GL","GNMA",IIf([ScheduleType]="MS","GNMA","FNMA")) AS Agency, 
   IIf([ScheduleType]="MS",18,IIf([ScheduleType]="GL",18,IIf([ScheduleType]="MA",2,IIf([ScheduleType]="AA",48,IIf([ScheduleType]="SS",18,18))))) AS Delay, 
   CurrentLoanBalance*ChicagoParticipation AS cBal, 
   OriginalAmount, ChicagoParticipation, l.CEFee, l.DeliveryCommitmentNumber, PIAmount, interestRate
FROM tblMPFLoansAggregateSource AS l LEFT JOIN [MPF Hedges] AS h ON l.LoanNumber = h.LoanNumber






MPFrepos
DM_LAS_Daily_LossGain_CH
DM_MPF_Delivery_Commitment_CH


midlleWare
M32c_x_tblBloombergFIXEDExport
MW_DataLoadStatistics
UV_QRM_MBS_base
Instrument_MBS_MPF
Portfolio_Contents_MBS_MPF


MPFFHLBDW
UV_AdjustementSubFactors_NOSHFD
UV_MasterCommitment_NOSHFD
UV_AllLoansToMPFLoansAggegateSourceOne_NOSHFD
UV_Code_NOSHFD
UV_DeliveryCommitment_NOSHFD
UV_loan_NOSHFD
UV_loanFunding_NOSHFD
UV_LoanFundingDateView_NOSHFD
UV_LoanFundingSumView_NOSHFD
UV_LoanRatesAgentFees_NOSHFD
UV_PFI_NOSHFD
UV_PFILoan_NOSHFD
UV_Product_NOSHFD
UV_Schedule_NOSHFD

/*******VIEW: UV_AllMPFLoans_NOSHFD**************************************************************************  
Date Created	: 01/24/2014  
Author			: Ajinkya Korade  
TFS				: 4329  
Item			: 1661
Version			: DataMart 11.0 HF 25
Purpose			: AIPhase 1 ::The AI-GNMA query was running slower when executed through Access db,
				  so a user view created for fast retrieval of Data.      

**********************************************************************************************************************/
CREATE VIEW UV_AllMPFLoans_NOSHFD
AS 
	SELECT	DWView.City,
			DWView.State,
			DWView.PFINumber,
			DWView.MANumber,
			DWView.DeliveryCommitmentNumber,
			DWView.LoanNumber,
			DWView.LoanRecordCreationDate,
			DWView.LastFundingEntryDate,
			DWView.ClosingDate,
			DWView.Amount AS MPFBalance,
			DWView.LoanAmount,
			DWView.TransactionCode,
			DWView.InterestRate,
			DWView.ProductCode,
			DWView.NumberOfMonths,
			DWView.FirstPaymentDate,
			DWView.MaturityDate,
			DWView.PIAmount,
			DWView.ScheduleCode,
			DWView.NonBusDayTrnEffDate,
			DWView.RemittanceTypeID,
			DWView.ExcessServicingFee,
			DWView.ServicingFee,
			DWView.BatchID,
			DWView.ParticipationOrgKey,
			DWView.ParticipationPercent,
			DWView.LoanToValue AS OrigLTV,
			LAS.LOANLoanInvestmentStatus
	FROM	LOS_PROD.dbo.UV_AllMPFLoansSubView_NOSHFD	DWView
			INNER JOIN dbo.UV_LAS_Loan	AS LAS
			ON DWView.LoanNumber = LAS.LoanNumber
	WHERE	LAS.LOANLoanInvestmentStatus NOT IN (06,03)

/*
***** VIEW:  UV_HFSLoans_NOSHFD  *************************************************************************************************************************************************
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
* Version Number		Date			Author   			Description of Change:
* ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
* AI Iteration 2		01/08/2014		Sahitya Bhandarkar	TFS#4314 :: Item#1613 - Created view for LAS HFS loans to improve performance of commonly run Access query by FHLBC on its environment for MPFFives HFS Loans
* Datamart 11.0	HF 24	01/22/2014		Korade Ajinkya		TFS#4328 :: item#1658 :: AIPhase 1 :: Added a distinct clause to avoid duplicate records.
* AI Iteration 2		01/29/2014		Avdhut Vaidya		TFS#4328 :: item#1658 :: Ai Phase 1 :: using an 'iNNER JOiN' instead of 'LEFT JOiN' while
																		joining UV_DCParticipationView since otherwise it gives records of ChicagoParticipation = 0
********************************************************************************************************************************************************************************************
*/

CREATE VIEW [dbo].[UV_HFSLoans_NOSHFD]
AS
SELECT DISTINCT lf.LoanNumber
	, LASHFSLoans.LOANLoaninvestmentstatus AS LoanInvestmentStatus
	, ISNULL(dcp.ParticipationPercent, 0) AS ChicagoParticipation
	, lf.InterestRate
	, lf.FundingDate
	, s.ProductCode
	, DMMPFDL.PHActionCode AS Action_Code
FROM LOS_Prod.dbo.LoanFunding lf
	INNER JOIN LOS_prod.dbo.DeliveryCommitment dc ON dc.DeliveryCommitmentID = lf.DeliveryCommitmentNumber
	INNER JOIN LOS_prod.dbo.Schedule s ON s.ScheduleCode = dc.ScheduleCode
	--item#1658 :: Avdhut Vaidya :: changed the LEFT JOiN to iNNER JOiN
	iNNER JOIN LOS_prod.dbo.UV_DCParticipationView dcp ON dcp.DeliveryCommitmentNumber = dc.DeliveryCommitmentID
													   AND dcp.ParticipationOrgKey = 3
	INNER JOIN MPFFHLBDW.dbo.UV_LAS_HFSLoans LASHFSLoans ON LASHFSLoans.LoanNumber = lf.LoanNumber
	INNER JOIN MPFFHLBDW.dbo.DM_MPF_Daily_Loan DMMPFDL ON DMMPFDL.LoanNumber = lf.LoanNumber
	INNER JOIN LOS_prod.dbo.MasterCommitment mc ON dc.MasterCommitmentKey = mc.MasterCommitmentKey
	INNER JOIN LOS_prod.dbo.MCModelType mcmt ON mc.MCModelTypeKey = mcmt.MCModelTypeKey
WHERE (s.ProductCode) in ('GL03', 'GL05')
AND mcmt.ServicingModelKey NOT IN (2,4)

