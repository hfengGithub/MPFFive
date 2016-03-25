-- output: K:\MRO\Archive\yyyy\yyyy-mm\yyyy-mm-dd\Pricing\Pricing-HFV-Loans.csv
--         TPODSLoc
--		polypaths Input	---==== create pf file
--		Instrument_MBS_MPF ==== to MiddleWare
--		Portfolio_Contents_MBS_MPF ==== to MiddleWare


---- ================================================== 20140903 GNII (HFS) BEGIN
------- qryInsertHFSLoansEOD   
-- Hua 20150417 ----- dbo_UV_HFSLoans_NOSHFD is loaded EOD, so we need both dbo_UV_HFSLoans_NOSHFD and tblMPFLoansAggregateSource. 
-- Hua 20150603 --  changed ScheduleType for the second part; first part will be updated when RemittanceTypeID is added in dbo_UV_HFSLoans_NOSHFD.
--                  put prefix "F" in cusip for FV loans
--		IIf(mpf.LoanInvestmentStatus="12",CInt(mpf.InterestRate*200)/2,INT((mpf.InterestRate*200)+0.5)/2) AS PassThruRate, LATER
-- Hua 20150808 -- changed the 1st ScheduleType and replaced dbo_UV_HFSLoans_NOSHFD with the new dbo_UV_FairValueOptionLoans_NOSHFD
-- Hua 20150810 assume dbo_UV_FairValueOptionLoans_NOSHFD contains real time data, and "06" records included
-- Hua 20160122 replace tblMPFLoansAggregateSource with tblTempMPF
-- Hua 20160224 Changed productCode and Cusip for GNMBS jumbo loan
	SELECT 	f.LoanNumber,
		f.LoanInvestmentStatus,
		f.ChicagoParticipation,
		f.InterestRate,
		f.FundingDate,
		NZ(mpf.ProductCode, f.ProductCode) AS productCode,
		f.Action_Code,
		IIf(Left(ProductCode,2)="GL","GL",
			IIf(f.RemittanceTypeID=1,"SS",
			IIf(f.RemittanceTypeID=2,"AA", "MA"))) AS ScheduleType, 
		INT((f.InterestRate*200)+0.5)/2 AS PassThruRate,
		Right(ProductCode,2) AS AccountType,
		NZ(Year(mpf.LoanRecordCreationDate),Year(Date())) AS OriginationYear,
		IIf(f.LoanInvestmentStatus="12","F","") & (ScheduleType & OriginationYear & AccountType & PassThruRate*100) AS CUSIP,
		Cdbl(0) AS Price		
	INTO tblHFSLoansEOD
	FROM dbo_UV_FairValueOptionLoans_NOSHFD AS f LEFT JOIN tblTempMPF AS mpf ON f.LoanNumber = mpf.LoanNumber		
	WHERE f.LoanInvestmentStatus NOT IN ("03", "11")
	AND   (f.LOANSettlementDate is NULL OR Date()-1<f.LOANSettlementDate )


-------- qryAppendV2HFVloans ====== Hua 20150210 --- NO USE -- only as a holder
INSERT INTO tblHFSLoansEOD (loanNumber)
select LoanNumber 
from  [MPF Hedges]
where loanNumber<0

---- qryDeleteClosedHFSLoans
DELETE *
FROM tblHFSLoansEOD
WHERE (((tblHFSLoansEOD.LoanInvestmentStatus)="08") AND ((tblHFSLoansEOD.Action_Code) In ("60","65","70","71","72")));


-- ============================= rptHFSEODCheck ==================== BEGIN
----- qryCheckExistingHFSLoanCount
-- Hua 20150809 change dbo_UV_HFSLoans_NOSHFD to dbo_UV_FairValueOptionLoans_NOSHFD
-- Hua 20160224 Changed productCode checking for GNMBS jumbo loan
SELECT l.ProductCode, Count(l.LoanNumber) AS CountOfExitingLoans
FROM tblMPFLoansAggregateSource as l INNER JOIN dbo_UV_FairValueOptionLoans_NOSHFD as f ON l.LoanNumber = f.LoanNumber
GROUP BY l.ProductCode
HAVING (((l.ProductCode) In ("GL03","GL05","GL43","GL45")));

--------- in Query Builder
SELECT qryCheckExistingHFSLoanCount.ProductCode, qryCheckExistingHFSLoanCount.CountOfExitingLoans
FROM qryCheckExistingHFSLoanCount;

--- qryCheckNewHFSLoanCount
-- Hua 20160224 Changed to use tblHFSLoansEOD.productCode checking for GNMBS jumbo loan
SELECT tblHFSLoansEOD.ProductCode, Count(tblHFSLoansEOD.LoanNumber) AS CountOfLoanNumber
FROM tblMPFLoansAggregateSource RIGHT JOIN tblHFSLoansEOD ON tblMPFLoansAggregateSource.LoanNumber = tblHFSLoansEOD.LoanNumber
WHERE (((tblMPFLoansAggregateSource.LoanNumber) Is Null))
GROUP BY tblHFSLoansEOD.ProductCode, tblMPFLoansAggregateSource.ProductCode
HAVING (((tblHFSLoansEOD.ProductCode) In ("GL03","GL05","GL43","GL45")));

--------- in Query Builder
SELECT qryCheckNewHFSLoanCount.ProductCode, qryCheckNewHFSLoanCount.CountOfLoanNumber
FROM qryCheckNewHFSLoanCount;

--- qryCheckMissingHFSLoanCount
-- Hua 20160224 Changed productCode checking for GNMBS jumbo loan
SELECT tblMPFLoansAggregateSource.ProductCode, Count(*) AS CountOfMissingLoans
FROM tblMPFLoansAggregateSource LEFT JOIN tblHFSLoansEOD ON tblMPFLoansAggregateSource.LoanNumber = tblHFSLoansEOD.LoanNumber
WHERE tblHFSLoansEOD.LoanNumber Is Null
GROUP BY tblMPFLoansAggregateSource.ProductCode
HAVING (((tblMPFLoansAggregateSource.ProductCode) In ("GL03","GL05","GL43","GL45")));

SELECT qryCheckMissingHFSLoanCount.ProductCode, qryCheckMissingHFSLoanCount.CountOfMissingLoans
FROM qryCheckMissingHFSLoanCount;
-- ============================= rptHFSEODCheck ==================== END

--- qryUpdateHFSLoanPrice
UPDATE tblHFSLoansEOD INNER JOIN MPFPrice ON tblHFSLoansEOD.CUSIP = MPFPrice.CUSIP SET tblHFSLoansEOD.Price = MPFPrice.Price;

---- ============================================================== 20140903 GNII end


-- ======================================= 20140723 From MPFFHLBDW to tblMPFLoansAggregateSource ======== BEGIN
-- ========== load loan level data
-- DeleteForwardSettleDCs MPFHedges
DELETE [MPF Hedges].*, [MPF Hedges].HedgeID, [MPF Hedges].SwapHedgeID
FROM [MPF Hedges]
WHERE ((([MPF Hedges].HedgeID)="HedgeSA") AND (([MPF Hedges].SwapHedgeID)="HedgeSA"));
-- qryMaketbltempMA
-- qryDeletetblMPFLoansAggregateSource

--------- qryselectallloans ==== passThrough version
-- Hua 20150914 use uv_maxPayDate instead of tblnorwestloansforfives to improve the performance --- in PRODUCTION
-- Hua 20151223 use UV_FairValueOptionLoans_NOSHFD to filter out the sold GNMBS loans
-- Hua 20160221 productCode, Change New/OldCoupon for GNMBS jumbo loans =========================================================== need to manually update
SELECT City, State, mpf.PFINumber, mpf.MANumber, DeliveryCommitmentNumber, mpf.LoanNumber, LoanRecordCreationDate, LastFundingEntryDate, 
	ClosingDate, ISNULL(m.ScheduleEndPrincipalBal, mpf.[MPFBalance]) AS MPFBalance, LoanAmount, TransactionCode, mpf.InterestRate, 
	CASE WHEN mpf.DeliveryCommitmentNumber IN (587414) THEN 'GL4'+RIGHT(mpf.ProductCode,1) ELSE mpf.ProductCode END AS ProductCode, NumberOfMonths, 
	CASE 
	WHEN Left(mpf.productCode,2) = 'GL' THEN 
		CASE WHEN substring(mpf.productCode,3,1) IN ('0','4') THEN ROUND(mpf.InterestRate*200,0)/200-.005 ELSE mpf.InterestRate - 0.0046 END
	WHEN mpf.productCode = 'FX15' THEN mpf.InterestRate - 0.0038
	ELSE mpf.InterestRate - 0.0039 END AS NewCoupon, 
	CASE 
	WHEN Left(mpf.productCode,2) = 'GL' THEN 
		CASE WHEN substring(mpf.productCode,3,1) IN ('0','4') THEN ROUND(mpf.InterestRate*200,0)/200-.005 ELSE mpf.InterestRate - (ServicingFee + ExcessServicingFee) END
	ELSE mpf.InterestRate - 0.0039 END AS CouponOld, 
	ROUND(mpf.InterestRate * 800,0) / 800 AS PrepaymentInterestRate, FirstPaymentDate, MaturityDate, PIAmount, ScheduleCode, NonBusDayTrnEffDate, 
	mpf.RemittanceTypeID, ExcessServicingFee, ServicingFee, BatchID, ParticipationOrgKey, ParticipationPercent AS ChicagoParticipation, 
	OrigLTV, LOANLoanInvestmentStatus
FROM (UV_AllMPFLoans_NOSHFD AS mpf LEFT JOIN uv_maxPayDate as m ON mpf.LoanNumber = m.LoanNumber) 
LEFT JOIN dbo.UV_FairValueOptionLoans_NOSHFD AS f ON mpf.loanNumber=f.loanNumber
WHERE mpf.MPFBalance>0
and   (m.loanNumber is null or m.ScheduleEndPrincipalBal>0)
and   mpf.ParticipationOrgKey = 3
and   mpf.LOANLoanInvestmentStatus <> '09'
and   (mpf.LOANLoanInvestmentStatus <> '06' OR f.loanNumber IS NOT NULL)


-- =============COUPON:
Loans: 
GN MBS: INT((InterestRate*200)-0.5)/200
GL: [InterestRate] - 0.0046 (-.0044 for loans originated later than 2/1/2007)
FX15: [InterestRate] - 0.0038
FX20, FX30: [InterestRate] - 0.0039


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
-- Hua 20150302 Added LoanInvestmentStatus, alter table tblMPFLoansAggregatesource ADD LoanInvestmentStatus
-- Hua 20150313 Changed newCoupon2 by 2bps
-- Hua 20160122 drop mpfBalance>0; filter out LOANLoanInvestmentStatus="06"
-- Hua 20160221 change the newcoupon2 to handle the jumbo GNMBS (temprarilly)
SELECT m.city, m.state, m.pfinumber, m.manumber, m.deliverycommitmentnumber, m.loannumber, m.loanrecordcreationdate, m.lastfundingentrydate, m.closingdate, m.mpfbalance, 
	m.loanamount, m.transactioncode, m.interestrate, m.ProductCode, m.prepaymentinterestrate, m.firstpaymentdate, m.maturitydate, 
	
	IIf(productCode IN ("GL15","GL30"), IIf(ma.entryDate>=#2/1/2007#,m.newcoupon+0.0002,m.newcoupon), m.newcoupon) AS newcoupon2, 
	
	m.couponold, 
	m.piamount, m.numberofmonths, m.schedulecode, m.nonbusdaytrneffdate, m.remittancetypeid, tblmonthenddate.monthenddate, m.excessservicingfee, m.servicingfee, 
	[cefee] + [ceperformance] AS macefee, ma.ceperformance AS ceperformancefee, m.batchid, m.chicagoparticipation, 
	m.OrigLTV, LOANLoanInvestmentStatus as LoanInvestmentStatus
FROM tblmonthenddate, tbltempmpf as m INNER JOIN tbltempma as ma ON (m.manumber = ma.manumber) AND (m.pfinumber = ma.pfinumber)
WHERE m.LOANLoanInvestmentStatus<>"06"
AND  ((m.nonbusdaytrneffdate) <= [tblmonthenddate]![monthenddate])



------ Hua 20150728 temporarilly in PRODUCTION to filter out the dups from dbo_UV_AllMPFLoans_NOSHFD
FROM tblmonthenddate, (select distinct * from tbltempmpf) as m INNER JOIN tbltempma as ma ON (m.manumber = ma.manumber) AND (m.pfinumber = ma.pfinumber)

--- qryAppendAllLoansStepTwo 
-- Hua 20150302 Added LoanInvestmentStatus, changed currentLoanBalance
-- Hua 20150427 Drop tblnorwestloansforfives. See qryselectallloans 20150423 change log
-- Hua 20150626 put in PRODUCTION
INSERT INTO tblMPFLoansAggregatesource ( loannumber, deliverycommitmentnumber, manumber, pfinumber, loanrecordcreationdate, lastfundingentrydate, 
	originalloanclosingdate, mpfbalance, originalamount, transactioncode, interestrate, coupon, couponold, prepaymentinterestrate, 
	firstpaymentdate, maturitydate, piamount, numberofmonths, schedulecode, productcode, portfolioindicator, remittancetypeid, age, 
	chicagoparticipation, currentloanbalance, entrydate, excessservicingfee, servicingfee, cefee, ceperformancefee, OrigLTV, LoanInvestmentStatus )
SELECT a.loannumber, a.deliverycommitmentnumber, a.manumber, a.pfinumber, a.loanrecordcreationdate, a.lastfundingentrydate, 
	a.closingdate AS originalloanclosingdate, a.mpfbalance, a.loanamount, a.transactioncode, a.interestrate, a.newcoupon2, a.couponold, 
	a.prepaymentinterestrate, a.firstpaymentdate, a.maturitydate, a.piamount, a.numberofmonths, a.schedulecode, a.productcode, 
	Iif([batchid] IS NOT NULL,"BATCH","FLOW") AS portfolioindicator, a.remittancetypeid, 
	([numberofmonths]-Iif(Datediff("m",a.monthenddate,maturitydate)>numberofmonths,numberofmonths, Datediff("m",a.monthenddate,maturitydate))) AS age, 
	a.chicagoparticipation, mpfbalance AS currentloanbalance, 
	a.monthenddate, a.excessservicingfee, a.servicingfee, a.macefee, a.ceperformancefee, a.OrigLTV, LoanInvestmentStatus
FROM qryselectallloanssteptwo as a 
WHERE ((a.chicagoparticipation) > 0)


-- qryMaketbltempMA – 
-- Hua 20150313 Added entryDate field
SELECT PFINumber, MANumber, CEFee, ProgramCode, IIf([CEPerformanceFee] Is Null,0,[CEPerformanceFee]) AS CEPerformance, entryDate 
INTO tblTempMA
FROM tblMasterAgreement;

======================================= 20140723 From MPFDW to tblMPFLoansAggregateSource ========= End

	
-------- the difference between MPFBalance and currentLoanBalance
SELECT count(a.LoanNumber) as cnt,  sum(s.CurrentLoanBalance*s.ChicagoParticipation)/1000000 as cbal, sum(a.MPFBalance*ParticipationPercent)/1000000 as amta, 
sum(s.MPFBalance*ChicagoParticipation)/1000000 as amts
FROM dbo_UV_AllMPFLoans_NOSHFD as a INNER JOIN tblMPFLoansAggregateSource1 as s ON a.LoanNumber = s.LoanNumber
WHERE (((a.ParticipationOrgKey)=3) and LOANLoanInvestmentStatus="01")

cnt	cbal	amta	amts
150555	7537.678116127	12373.0520800772	12374.0920751473

-- ================================================= Forward DC ============== BEGIN
-- ======== Load Chicago DC"s
-- qryFundedandUnfundedDCsbyPFI 
--  20141024 Hua changed CUSIP
-- 20150227 Hua Changed Delay, Cusip. v2 cusip same as v1, date tell difference
-- 20150317 Hua Changed the coup for GL +2bps
-- Hua 20150629 changed (InterestRate- 0.0035) to (InterestRate- 0.0038) for FX15
-- SELECT ProductCode, 
-- 	IIf(Left([ProductCode],2)="GL",IIf(Right([ProductCode],2)="03" OR Right([ProductCode],2)="05","GNMA2","GNMA"),
-- 		IIf(Left([ProductCode],2)="FX","FNMA"," ")) AS Agency, 
-- 	NoteRate, Fee, CDbl(0) AS Price, DeliveryStatus, ScheduleType, 
-- 	IIf(Left([ProductCode],2)="GL",1,dbo_UV_DCParticipation_MRA_NOSHFD.RemittanceTypeID) AS RemittanceTypeID, 
-- 	IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=2,"AA","MA"))) AS [Schedule Type], 
-- 	IIf([Schedule Type]="MA",2,IIf([Schedule Type]="AA",48,18)) AS Delay, 
-- 	EntryDate, EntryTime, DeliveryDate, 1 AS DTF, FullName, "1" AS PPMult, PFINumber, MANumber, DeliveryCommitmentNumber, DeliveryAmount, 
-- 	nz([DeliveryAmount])*nz([Participation]) AS DeliveryAmountP, FundedAmount, nz([FundedAmount])*nz([Participation]) AS FundedAmountP, 
-- 	([DeliveryAmount])-([FundedAmount]) AS UnfundedAmount, [DeliveryAmountP]-[FundedAmountP] AS UnfundedAmountP, LastUpdatedDate, 
-- 	ScheduleCode, Participation, 
-- 	Year([DeliveryDate]) AS DeliveryYear, Month([DeliveryDate]) AS DeliveryMonth, Day([DeliveryDate]) AS DeliveryDay, 
-- 	[Schedule Type] & [DeliveryYear] & IIf([DeliveryMonth]<10,"0","") & CStr([DeliveryMonth]) & Right([ProductCode],2) & IIF(
-- 		Right([ProductCode],2) = "03" OR Right([ProductCode],2) = "05", INT(([NoteRate]*200)+0.5)/2, (CInt([NoteRate]*200)/2))*1000 AS CUSIP, 
-- 	IsExtended, ServicingFee, ExcessServicingFee, CEFee, CEPerformanceFee,
-- 	IIf(Left([ProductCode],2)="GL",IIF(Right([ProductCode],2) = "03" OR Right([ProductCode],2) = "05", INT((NoteRate*200)-0.5)/200, CDbl([NoteRate])-0.0044),
-- 		IIf([ProductCode]="FX15",CDbl([NoteRate])-0.0038,CDbl([NoteRate])-0.0039)) AS Coup, 
-- 	IIf(Left([ProductCode],2)="GL",IIF(Right([ProductCode],2) = "03" OR Right([ProductCode],2) = "05", INT((NoteRate*200)-0.5)/200,
-- 		CDbl([NoteRate])-([ServicingFee]+[ExcessServicingFee]+[CEFee])),CDbl([NoteRate])-0.0039) AS CoupOld 
-- INTO ForwardSettleDCs
-- FROM dbo_UV_DCParticipation_MRA_NOSHFD  
-- WHERE ParticipationOrgKey = 3
-- ORDER BY NoteRate, DeliveryDate


-- qryFundedandUnfundedDCsbyPFI ------- need to be changed : FOR given DCnumber.. 20160204 ============================== need to be manually changed (DCnumbers)
-- Hua 20160210 Added AccountType "43", "45" and DCnumber holder to handle GNMBS jumbo; ProductCode, Agency, CUSIP, Coup, CoupOld
SELECT Iif(DeliveryCommitmentNumber IN (587414), "GL4" & right(v.ProductCode,1), v.ProductCode) AS ProductCode,
	IIf(Left(ProductCode,2)="GL",IIf(Right(ProductCode,2) IN ("03","05","43","45"),"GNMA2","GNMA"),
		IIf(Left(ProductCode,2)="FX","FNMA"," ")) AS Agency, 
	NoteRate, Fee, CDbl(0) AS Price, DeliveryStatus, ScheduleType, 
	IIf(Left(ProductCode,2)="GL",1,v.RemittanceTypeID) AS RemittanceTypeID, 
	IIf(Left(ProductCode,2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=2,"AA","MA"))) AS [Schedule Type], 
	IIf([Schedule Type]="MA",2,IIf([Schedule Type]="AA",48,18)) AS Delay, 
	EntryDate, EntryTime, DeliveryDate, 1 AS DTF, FullName, "1" AS PPMult, PFINumber, MANumber, DeliveryCommitmentNumber, DeliveryAmount, 
	nz([DeliveryAmount])*nz([Participation]) AS DeliveryAmountP, FundedAmount, nz([FundedAmount])*nz([Participation]) AS FundedAmountP, 
	([DeliveryAmount])-([FundedAmount]) AS UnfundedAmount, [DeliveryAmountP]-[FundedAmountP] AS UnfundedAmountP, LastUpdatedDate, 
	ScheduleCode, Participation, 
	Year([DeliveryDate]) AS DeliveryYear, Month([DeliveryDate]) AS DeliveryMonth, Day([DeliveryDate]) AS DeliveryDay, 
	
	[Schedule Type] & [DeliveryYear] & FORMAT(DeliveryMonth,"00") & Right(ProductCode,2) & IIF(
		Right(ProductCode,2) IN ("03","05","43","45"), INT(([NoteRate]*200)+0.5)/2, (CInt([NoteRate]*200)/2))*1000 AS CUSIP, 
		
	IsExtended, ServicingFee, ExcessServicingFee, CEFee, CEPerformanceFee,
	IIf(Left(ProductCode,2)="GL",IIF(Right(ProductCode,2) IN ("03","05","43","45"), INT((NoteRate*200)-0.5)/200, CDbl([NoteRate])-0.0044),
		IIf(ProductCode="FX15",CDbl([NoteRate])-0.0038,CDbl([NoteRate])-0.0039)) AS Coup, 
	IIf(Left(ProductCode,2)="GL",IIF(Right(ProductCode,2) IN ("03","05","43","45"), INT((NoteRate*200)-0.5)/200,
		CDbl([NoteRate])-([ServicingFee]+[ExcessServicingFee]+[CEFee])),CDbl([NoteRate])-0.0039) AS CoupOld 
INTO ForwardSettleDCs
FROM dbo_UV_DCParticipation_MRA_NOSHFD AS v
WHERE ParticipationOrgKey = 3
ORDER BY NoteRate, DeliveryDate



-- AppendForwardSettleDCs MPFHedges
INSERT INTO [MPF Hedges] ( DeliveryCommitmentNumber, Cusip, HedgeID, SwapHedgeID )
SELECT ForwardSettleDCs.DeliveryCommitmentNumber, "HC" & f.[DeliveryCommitmentNumber] AS Cusip, "HedgeSA" AS HedgeID, "HedgeSA" AS SwapHedgeID
FROM ForwardSettleDCs;

-- UpdateMPFForward15&20PPMult
UPDATE tblForwardCommitmentPalmsSource SET tblForwardCommitmentPalmsSource.PPMult = 1.3
WHERE (((tblForwardCommitmentPalmsSource.ProductCode) Like "FX15" Or (tblForwardCommitmentPalmsSource.ProductCode)="GL15" 
Or (tblForwardCommitmentPalmsSource.ProductCode)="FX20" Or (tblForwardCommitmentPalmsSource.ProductCode)="GL20") AND ((CDbl([tblForwardCommitmentPalmsSource].[Wac]))>=6.02));

-- UpdateMPFForward30PPMult
UPDATE tblForwardCommitmentPalmsSource SET tblForwardCommitmentPalmsSource.PPMult = 1.2
WHERE (((tblForwardCommitmentPalmsSource.ProductCode) Like "FX30" Or (tblForwardCommitmentPalmsSource.ProductCode)="GL30") AND ((CDbl([tblForwardCommitmentPalmsSource].[Wac]))>=7.04));

-- qryFundedandUnfundedDCsbyPFI Unhedged
SELECT ForwardSettleDCs.*, [MPF Hedges].DeliveryCommitmentNumber AS DCNHedged, [MPF Hedges].HedgeID
FROM ForwardSettleDCs LEFT JOIN [MPF Hedges] ON ForwardSettleDCs.DeliveryCommitmentNumber = [MPF Hedges].DeliveryCommitmentNumber
WHERE ((([MPF Hedges].HedgeID) Is Null));

-- qryFundedandUnfundedDCsbyPFI Hedged
SELECT ForwardSettleDCs.*, [MPF Hedges].DeliveryCommitmentNumber AS DCNHedged, [MPF Hedges].PurchaseGroup, [MPF Hedges].HedgeID, [MPF Hedges].SwapHedgeID
FROM ForwardSettleDCs LEFT JOIN [MPF Hedges] ON ForwardSettleDCs.DeliveryCommitmentNumber = [MPF Hedges].DeliveryCommitmentNumber
WHERE ((([MPF Hedges].HedgeID) Is Not Null));


-- ============ Aggregate MPF Forwards 201404

-- qryAggregateForwardSettleCommitments
-- Hua 20150227 Changed Cusip, CUSIPHedged
-- Hua 20150316 Changed Coup, CoupOld, prepCoup for GL loans later than 2/1/2007 
-- Hua 20150511 Changed WAC for GN MBS loans
-- Hua 20150602 fixed error by *100	
-- Hua 20150611 Added dbo_uv_masterCommitment_noshfd, put prefix "C" to CUSIP for MPFv2	
-- Hua 20150626 change the WAC from 
-- 		IIF(Mid([ProductCode],3,2) in ("03", "05"),INT(([NoteRate]*200)+0.5)/2, CInt([NoteRate]*200)/2) AS Wac 
-- 	to	100*SUM(UnfundedAmountP*NoteRate)/Notional AS Wac
--	changed [Account Type] from Wac as [Account Type], to RIGHT(f.CUSIP,4)/1000 AS [Account Type]
-- INSERT INTO tblForwardCommitmentPalmsSource ( 
-- 	ProductCode, RemittanceTypeID, DeliveryYear, Wac, DeliveryMonth, WASettleDay, ScheduleType, ScheduleType2, Delay, Portfolio, CUSIP, 
-- 	CusipHedged, Owam, [Sub Account Type], [Account Type], [Account Class], H1, H2, Notional, Mult, Factor, [Add Accrued?], PV, Swap, 
-- 	Age, Wam, Swam, [P/O], OAS, P, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSens, Lag, Lock, PPCRShift, PPFq, NxtPP, NxtRst1, 
-- 	PrepWac, Coup, CoupOld, PrepCoup, FA1, Const1, Mult1, Rate1, RF1, Floor1, Cap1, PF1, PC1, AB1, NxtRst2, BOOK, 
-- 	ClientName, AggNotional, WAPrice, NumberofCommitments, Agency )
-- SELECT ProductCode, f.RemittanceTypeID, DeliveryYear, 
-- 	100*SUM(UnfundedAmountP*NoteRate)/Notional AS Wac, 
-- 	DeliveryMonth, CInt(Sum(Abs([UnfundedAmountP])*[DeliveryDay])/Sum(Abs([UnfundedAmountP]))) AS WASettleDay, [Schedule Type], [Schedule Type], 
-- 	Last(Delay) AS LastOfDelay, "MPFForward" AS Portfolio, 	
-- 	IIF(m.investmentOption="04","C","") & f.CUSIP AS CUSIP, 
-- 	CUSIP AS CUSIPHedged, 	
-- 	IIF(Mid([ProductCode],3,2)="03",30,IIF(Mid([ProductCode],3,2)="05",15,Mid([ProductCode],3,2)))*12 AS Owam, 
-- 	ProductCode AS [Sub Account Type], 
-- 	
-- 	RIGHT(f.CUSIP,4)/1000 AS [Account Type], CStr([DeliveryYear]) AS [Account Class], 
-- 	CStr([WAPrice]) AS H1, CStr([WASettleDay]) AS H2, 
-- 	CStr(Sum([UnfundedAmountP])) AS Notional, "1" AS Mult, "1" AS Factor, "1" AS [Add Accrued?], "Mid" AS PV, "Mid" AS Swap, "1" AS Age, [OWAM]-1 AS Wam, 
-- 	IIf([OWAM]=180,160,IIf([OWAM]=360,335,IIf([OWAM]=240,220,""))) AS Swam, "OAS" AS [P/O], 0 AS OAS, 
-- 	Sum([deliveryamount]*[price])/Sum([deliveryamount]) AS P, 
-- 	IIf(DeliveryMonth<10,"0","") & CStr(DeliveryMonth) & "/" & IIf(WASettleDay<10,"0","") & CStr([WASettleDay]) & "/" & CStr(DeliveryYear) AS Settle, 
-- 	DLookUp("[RepoRate]","[tblPALMSRepo]") AS Repo, NextDateFwd([DeliveryYear],[DeliveryMonth],[WASettleDay],[Schedule Type]) AS NxtPmt, 
-- 	"AFT" AS PP, "0" AS PPConst, Last(PPMult) AS LastOfPPMult, "1" AS PPSens, "0" AS Lag, "0" AS Lock, 0 AS PPCRShift, "12" AS PPFq, 
-- 	[NxtPmt] AS NxtPP, [NxtPmt] AS NxtRst1, [Wac] AS PrepWac, 	
-- 
-- 	100*SUM(UnfundedAmountP*f.Coup)/SUM(UnfundedAmountP) AS Coup, 
-- 	100*SUM(UnfundedAmountP*f.CoupOld)/SUM(UnfundedAmountP) AS CoupOld, 
-- 	Coup AS PrepCoup, 
-- 	
-- 	"F" AS FA1, "0" AS Const1, "0" AS Mult1, "3L" AS Rate1, "12" AS RF1, "None" AS Floor1, "None" AS Cap1, "None" AS PF1, "None" AS PC1, "Bond" AS AB1, 
-- 	[NxtRst1] AS NxtRst2, CStr(Sum([UnfundedAmountP])) AS BOOK, "MPFForward" AS ClientName, Sum(UnfundedAmountP) AS AggNotional, 
-- 	0 AS WAPrice, Count(NoteRate) AS NumberofCommitments, Agency
-- FROM [qryFundedandUnfundedDCsbyPFI Unhedged] as f INNER JOIN dbo_uv_masterCommitment_noshfd as m on f.MAnumber=m.MAnumber
-- GROUP BY Agency, ProductCode, f.RemittanceTypeID, DeliveryYear, DeliveryMonth, [Schedule Type], IIF(m.investmentOption="04","C","") & f.CUSIP, f.cusip
-- HAVING Sum([UnfundedAmountP])>0;

-- qryAggregateForwardSettleCommitments
-- Hua 20150227 Changed Cusip, CUSIPHedged
-- Hua 20150316 Changed Coup, CoupOld, prepCoup for GL loans later than 2/1/2007 
-- Hua 20150511 Changed WAC for GN MBS loans
-- Hua 20150602 fixed error by *100	
-- Hua 20150611 Added dbo_uv_masterCommitment_noshfd, put prefix "C" to CUSIP for MPFv2	
-- Hua 20150626 change the WAC from 
-- 		IIF(Mid([ProductCode],3,2) in ("03", "05"),INT(([NoteRate]*200)+0.5)/2, CInt([NoteRate]*200)/2) AS Wac 
-- 	to	100*SUM(UnfundedAmountP*NoteRate)/Notional AS Wac
-- Hua 20160212 update Owam, dbo_uv_dm_mpf_daily_masterCommitments for dbo_uv_masterCommitment_noshfd
-- Hua 20160316 updated P (price) to use UnfundedAmountP as weight
INSERT INTO tblForwardCommitmentPalmsSource ( 
	ProductCode, RemittanceTypeID, DeliveryYear, Wac, DeliveryMonth, WASettleDay, ScheduleType, ScheduleType2, Delay, Portfolio, CUSIP, 
	CusipHedged, Owam, [Sub Account Type], [Account Type], [Account Class], H1, H2, Notional, Mult, Factor, [Add Accrued?], PV, Swap, 
	Age, Wam, Swam, [P/O], OAS, P, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSens, Lag, Lock, PPCRShift, PPFq, NxtPP, NxtRst1, 
	PrepWac, Coup, CoupOld, PrepCoup, FA1, Const1, Mult1, Rate1, RF1, Floor1, Cap1, PF1, PC1, AB1, NxtRst2, BOOK, 
	ClientName, AggNotional, WAPrice, NumberofCommitments, Agency )
SELECT ProductCode, f.RemittanceTypeID, DeliveryYear, 
	100*SUM(UnfundedAmountP*NoteRate)/Notional AS Wac, 
	DeliveryMonth, CInt(Sum(Abs([UnfundedAmountP])*[DeliveryDay])/Sum(Abs([UnfundedAmountP]))) AS WASettleDay, [Schedule Type], [Schedule Type], 
	Last(Delay) AS LastOfDelay, "MPFForward" AS Portfolio, 	
	IIF(m.investmentOption="04","C","") & f.CUSIP AS CUSIP, 
	CUSIP AS CUSIPHedged, 	
	IIF(Right(ProductCode,2) IN ("03","43"), 30, IIF(Right(ProductCode,2) IN ("05","45"), 15, Right(ProductCode,2) ))*12 AS Owam,
	ProductCode AS [Sub Account Type], 	
	RIGHT(f.CUSIP,4)/1000 AS [Account Type], CStr([DeliveryYear]) AS [Account Class], 
	CStr([WAPrice]) AS H1, CStr([WASettleDay]) AS H2, 
	CStr(Sum([UnfundedAmountP])) AS Notional, "1" AS Mult, "1" AS Factor, "1" AS [Add Accrued?], "Mid" AS PV, "Mid" AS Swap, "1" AS Age, [OWAM]-1 AS Wam, 
	IIf([OWAM]=180,160,IIf([OWAM]=360,335,IIf([OWAM]=240,220,""))) AS Swam, "OAS" AS [P/O], 0 AS OAS, 
	Sum([UnfundedAmountP]*[price])/Notional AS P, 
	IIf(DeliveryMonth<10,"0","") & CStr(DeliveryMonth) & "/" & IIf(WASettleDay<10,"0","") & CStr([WASettleDay]) & "/" & CStr(DeliveryYear) AS Settle, 
	DLookUp("[RepoRate]","[tblPALMSRepo]") AS Repo, NextDateFwd([DeliveryYear],[DeliveryMonth],[WASettleDay],[Schedule Type]) AS NxtPmt, 
	"AFT" AS PP, "0" AS PPConst, Last(PPMult) AS LastOfPPMult, "1" AS PPSens, "0" AS Lag, "0" AS Lock, 0 AS PPCRShift, "12" AS PPFq, 
	[NxtPmt] AS NxtPP, [NxtPmt] AS NxtRst1, [Wac] AS PrepWac, 	
	100*SUM(UnfundedAmountP*f.Coup)/SUM(UnfundedAmountP) AS Coup, 
	100*SUM(UnfundedAmountP*f.CoupOld)/SUM(UnfundedAmountP) AS CoupOld, 
	Coup AS PrepCoup, 	
	"F" AS FA1, "0" AS Const1, "0" AS Mult1, "3L" AS Rate1, "12" AS RF1, "None" AS Floor1, "None" AS Cap1, "None" AS PF1, "None" AS PC1, "Bond" AS AB1, 
	[NxtRst1] AS NxtRst2, CStr(Sum([UnfundedAmountP])) AS BOOK, "MPFForward" AS ClientName, Sum(UnfundedAmountP) AS AggNotional, 
	0 AS WAPrice, Count(NoteRate) AS NumberofCommitments, Agency
FROM [qryFundedandUnfundedDCsbyPFI Unhedged] as f INNER JOIN dbo_uv_dm_mpf_daily_masterCommitments as m on f.MAnumber=m.MAnumber
GROUP BY Agency, ProductCode, f.RemittanceTypeID, DeliveryYear, DeliveryMonth, [Schedule Type], IIF(m.investmentOption="04","C","") & f.CUSIP, f.cusip
HAVING Sum([UnfundedAmountP])>0;

-- ================================================= Forward DC ============== END


InvestmentOption="04"

------- qryDeleteMPFandForwardsfromInstrumenttblProduction --- Command59_Click() Not used
DELETE [Instrument-Palms].CUSIP AS Expr1, [Instrument-Palms].BondCUSIP AS Expr2, [Instrument-Palms].[CUSIP], [Instrument-Palms].[BondCUSIP]
FROM [Instrument-Palms]
WHERE ((([Instrument-Palms].[CUSIP]) Like "AA*" Or ([Instrument-Palms].[CUSIP]) Like "GL*" Or ([Instrument-Palms].[CUSIP]) Like "MS*" Or ([Instrument-Palms].[CUSIP]) Like "SS*" Or ([Instrument-Palms].[CUSIP]) Like "DC*" Or ([Instrument-Palms].[CUSIP]) Like "HC*" Or ([Instrument-Palms].[CUSIP]) Like "MA*") AND (([Instrument-Palms].[BondCUSIP])="MPF" Or ([Instrument-Palms].[BondCUSIP])="MPFForward"));


-- =================================================================== Aggregate MPF loans ======== BEGIN
-- qryUpdatePricesIntblSecuritiesPalmsSource
UPDATE tblSecuritiesPalmsSource INNER JOIN tblFinalPriceSheet ON tblSecuritiesPalmsSource.CUSIP = tblFinalPriceSheet.Cusip SET tblSecuritiesPalmsSource.AggPrice = [New Price]
WHERE (((tblSecuritiesPalmsSource.H1)="Seasoned"));


------------ MPFForwardPrice and MPFPrice are spread sheets
-- qryMPFForwardPrice
SELECT CatType, ProdType, Year, Date, [MBS Coupon], IIf([MPF Rate] Is Null,Null,CStr([MPF Rate]/100)) AS Rate, Day, Price
FROM MPFForwardPrice
WHERE (((CatType) Is Not Null) AND ((ProdType) Is Not Null) AND ((Year) Is Not Null) AND ((Date) Is Not Null) AND (([MBS Coupon]) Is Not Null) AND ((IIf([MPF Rate] Is Null,Null,CStr([MPF Rate]/100))) Is Not Null) AND ((Price) Is Not Null));


--- ==================================================== Run Palm data report

---------- qryPalmsDataSource:  
-- Hua 20141208 CHANGED Delay AND PassThruRate
-- Hua 20150227 put "C" and "F" prefix in Cusip for v2 loans. temp: NumberOfMonths-age as WAM
-- Hua 20150812 make age>0
-- Hua 20150820 changed interestRate in tblMPFLoansAggregateSource and tblTempMPF from single to double
-- Hua 20160210 Changed AccountType ("43", "45"),  to handle GNMBS jumbo; OriginationMonth, AccountType, PassThruRate, Swam, Agency
SELECT "MPF" AS Portfolio, 
	IIf(Left([a.ProductCode],2)="MS","MS",IIf(Left([a.ProductCode],2)="GL","GL",IIf([a.RemittanceTypeID]=1,"SS",IIf([a.RemittanceTypeID]=3,"MA","AA")))) AS ScheduleType, 
	True AS [Active?], 
	IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]), Year([a.LoanRecordCreationDate])) AS OriginationYear, 
	a.LoanRecordCreationDate, a.LoanNumber, a.DeliveryCommitmentNumber, 
	IIf([PortfolioIndicator]="BATCH", FORMAT(Month(OriginalLoanClosingDate),"00"), FORMAT(Month([a.LoanRecordCreationDate]),"00")) AS OriginationMonth, 
	Right([a.ProductCode],2)  AS AccountType, 
	IIF(AccountType IN ("03","05","43","45"), INT((a.InterestRate*200)+0.5)/2, (CInt([a.InterestRate]*200)/2)) AS PassThruRate, 
	IIf(LoanInvestmentStatus="11","C",
		IIf(LoanInvestmentStatus="12","F","")) & ScheduleType & OriginationYear & AccountType & (PassThruRate*100) AS CUSIP, 
	[OriginationYear] AS AccountClass, [a.MPFBalance]*[a.ChicagoParticipation] AS Notional, 
	1 AS Mult, IIf([a.CurrentLoanBalance]=0,1,[a.CurrentLoanBalance]/[a.OriginalAmount]) AS Factor, 
	"1" AS [Add Accrued?], "Mid" AS PV, "Mid" AS Swap, (CInt([a.InterestRate]*100000)/100000)*100 AS Wac, (CInt([a.Coupon]*100000)/100000)*100 AS Coup, [Wac]-[Coup] AS Diff,
	IIf(a.age<2, numberofmonths-1, CInt(NPer(([a.InterestRate]/12),([a.PIAmount]*-1),[a.CurrentLoanBalance]))) AS WAM, 
	IIf(a.age<2, 1, a.age) as Age,	
	IIf(AccountType IN ("05","15","45"),160, IIf(AccountType="20",220,IIf([AccountType] IN ("03","43","30"),335,""))) AS Swam, 
	tblPALMSRepo.RepoRate AS Repo, 
	a.EntryDate AS Settle, NextDate([Settle],[ScheduleType]) AS NxtPmt, "0" AS PPConst, "1" AS PPMult, "1" AS PPSense, "0" AS Lag, "0" AS Lock, 
	"12" AS PPFq, [NxtPmt] AS NxtPP, NextDate([Settle],[ScheduleType]) AS NxtRst1, [Wac] AS PrepWac, [Coup] AS PrepCoup, "Fixed" AS FA, 
	"0" AS Const2, 1 AS Mult1, 0 AS Mult2, "3L" AS Rate, "12" AS RF, -1000000000 AS Floor, 1000000000 AS Cap, 1 AS PF, 1000000000 AS PF1, 
	1000000000 AS PF2, 1000000000 AS PC, "Bond" AS AB, [NxtRst1] AS NxtRst2, 0 AS LookBack2, 
	0 AS LookBackRate, EntryDate AS WamDate, "CPR" AS PPUnits, 0 AS PPCRShift, 0 AS RcvCurrEscr, 0 AS PayCurrEscr, "StraightLine" AS AmortMethod, 
	"MPFProgram" AS ClientName, a.PrepaymentInterestRate, a.CurrentLoanBalance AS AcctBalance, 
	IIf([ScheduleType]="GL", IIf(AccountType IN ("03","05","43","45"),"GNMA2","GNMA"),IIf([ScheduleType]="MS","GNMA","FNMA")) AS Agency,
	IIf([ScheduleType]="MS",18,IIf([ScheduleType]="GL",18,IIf([ScheduleType]="MA",2,IIf([ScheduleType]="AA",48, 18)))) AS Delay, 
	a.NumberOfMonths AS OWAM, 0 AS Ballon, 1 AS IF, 0 AS Const1, 0 AS [Int Coup Mult], 1 AS [PP Coup Mult], -10000000000 AS [Sum Floor], 
	10000000000 AS [Sum Cap], "None" AS [Servicing Model], "None" AS [Loss Model], 0 AS [Sched Cap?], a.CurrentLoanBalance, 
	a.OriginalAmount, a.ChicagoParticipation
FROM tblMPFLoansAggregateSource AS a, tblPALMSRepo

-- qryHedgedPalmsDataSource:
SELECT qryPalmsDataSource.*, 
   [MPF Hedges].DeliveryCommitmentNumber AS DeliveryCommitmentNumber, 
   CStr([qryPalmsDataSource].[DeliveryCommitmentNumber]) AS DCNText, 
   "DC" & IIF([SwapHedgeID] is Null, "00", Right([SwapHedgeID], 2)) & [ScheduleType] & Right([AccountClass],4) & [AccountType] & [PassThruRate] AS CusipHedged, 
   [MPF Hedges].PurchaseGroup, 
   [MPF Hedges].HedgeID
FROM qryPalmsDataSource inner JOIN [MPF Hedges] ON qryPalmsDataSource.LoanNumber = [MPF Hedges].LoanNumber
WHERE ((([MPF Hedges].HedgeID) Is Not Null));

-- qryUnhedgedPalmsDataSource --
SELECT qryPalmsDataSource.*, [MPF Hedges].DeliveryCommitmentNumber AS DeliveryCommitmentNumber, [MPF Hedges].HedgeID
FROM qryPalmsDataSource LEFT JOIN [MPF Hedges] ON qryPalmsDataSource.LoanNumber = [MPF Hedges].LoanNumber
WHERE ((([MPF Hedges].HedgeID) Is Null));

-- qryAppendUnhedgedDataTotblSecuritiesPalmsSource
-- Hua 20150619 Added group by cusip
INSERT INTO tblSecuritiesPalmsSource ( AccountClass, AccountType, ScheduleType, PassThruRate, Portfolio, CUSIP, H2, AggNotional, Mult, AggFactor, [Add Accrued?], PV, 
	Swap, AggWac, AggCoup, AggWam, AggAge, H1, Swam, AggOWAM, [P/O], AggPrice, OAS, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSense, Lag, Lock, PPFq, NxtPP, NxtRst1, 
	AggPrepWac, AggPrepCoup, FA, Const1, Const2, Rate, RF, Floor, Cap, PF, PF1, PF2, PC, AB, NxtRst2, LookBackRate, LookBack, WamDate, PPUnits, PPCRShift, RcvCurrEscr, 
	PayCurrEscr, AggBookPrice, AmortMethod, ClientName, [Active?], Agency, [Int Coup Mult], [PP Coup Mult], [Sum Floor], [Sum Cap], [Servicing Model], [Loss Model], 
	[Sched Cap?], Delay, Ballon, IF, Mult1, Mult2, PrepaymentInterestRate )
SELECT s.AccountClass, s.AccountType, s.ScheduleType, s.PassThruRate, Last(s.Portfolio) AS Portfolio, CUSIP, Count(s.LoanNumber) AS H2, 
	Sum(s.Notional) AS AggNotional, Last(s.Mult) AS Mult, 
	IIf(Sum([CurrentLoanBalance])=0,Sum([OriginalAmount])/Sum([OriginalAmount]),Sum([CurrentLoanBalance])/Sum([OriginalAmount])) AS AggFactor, 
	Last(s.[Add Accrued?]) AS [Add Accrued?], Last(s.PV) AS PV, Last(s.Swap) AS Swap, 
	(Sum([Wac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWac, 
	(Sum([Coup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggCoup, 
	CInt(Sum([Wam]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWam, 
	(Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggAge, 
	IIf(Val([AccountType])=30 And [AggWam]<340,"Seasoned",IIf(Val([AccountType])=20 And [AggWam]<220,"Seasoned",IIf(Val([AccountType])=15 And [AggWam]<160,"Seasoned","MPF"))) AS H1, 
	Last(s.Swam) AS Swam, 
	CInt(Sum([OWAM]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggOWAM, 
	"OAS" AS [P/O], 0 AS AggPrice, 0 AS OAS, Last(s.Settle) AS Settle, Last(s.Repo) AS Repo, Last(s.NxtPmt) AS NxtPmt, "AFT" AS PP, Last(s.PPConst) AS PPConst, 
	Last(s.PPMult) AS PPMult, Last(s.PPSense) AS PPSense, Last(s.Lag) AS Lag, Last(s.Lock) AS Lock, Last(s.PPFq) AS PPFq, Last(s.NxtPP) AS NxtPP, Last(s.NxtRst1) AS NxtRst1, 
	(Sum([PrepWac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepWac, 
	(Sum([PrepCoup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepCoup, 
	Last(s.FA) AS FA, Last(s.Const1) AS Const1, Last(s.Const2) AS Const2, Last(s.Rate) AS Rate, Last(s.RF) AS RF, Last(s.Floor) AS Floor, Last(s.Cap) AS Cap, Last(s.PF) AS PF, 
	Last(s.PF1) AS PF1, Last(s.PF2) AS PF2, Last(s.PC) AS PC, Last(s.AB) AS AB, Last(s.NxtRst2) AS NxtRst2, Last(s.LookBackRate) AS LookBackRate1, Last(s.LookBack2) AS LookBack2, 
	Last(CStr(s.[WamDate])) AS WamDate, Last(s.PPUnits) AS PPUnits, Avg(s.PPCRShift) AS PPCRShift, Last(s.RcvCurrEscr) AS RcvCurrEscr, 
	Last(s.PayCurrEscr) AS PayCurrEscr, 
	Sum([CurrentLoanBalance]*[ChicagoParticipation]) AS AggBookPrice, 
	Last(s.AmortMethod) AS AmortMethod, Last(s.ClientName) AS ClientName, Last(s.[Active?]) AS [Active?], Last(s.Agency) AS Agency, Last(s.[Int Coup Mult]) AS [Int Coup Mult], 
	Last(s.[PP Coup Mult]) AS [PP Coup Mult], Last(s.[Sum Floor]) AS [Sum Floor], Last(s.[Sum Cap]) AS [Sum Cap], Last(s.[Servicing Model]) AS [Servicing Model], 
	Last(s.[Loss Model]) AS [Loss Model], Last(s.[Sched Cap?]) AS [Sched Cap?], Last(s.Delay) AS Delay, Last(s.Ballon) AS Ballon, Last(s.IF) AS IF, Last(s.Mult1) AS Mult1, 
	Last(s.Mult2) AS Mult2, (Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepaymentInterestRate
FROM qryUnhedgedPalmsDataSource AS s
GROUP BY s.AccountClass, s.AccountType, s.ScheduleType, s.PassThruRate, s.OriginationYear, CUSIP
HAVING ((((Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])))>30))
ORDER BY s.OriginationYear, s.AccountType DESC , s.ScheduleType DESC , 
	(Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation]));
	

-- qryAppendUnhedgedDataTotblSecuritiesPalmsSource - NONSEASONED
-- Hua 20150619 Added group by cusip
INSERT INTO tblSecuritiesPalmsSource ( AccountClass, AccountType, ScheduleType, PassThruRate, Portfolio, CUSIP, H2, AggNotional, Mult, AggFactor, [Add Accrued?], PV, Swap, 
	AggWac, AggCoup, AggWam, AggAge, H1, Swam, AggOWAM, [P/O], AggPrice, OAS, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSense, Lag, Lock, PPFq, NxtPP, NxtRst1, AggPrepWac, 
	AggPrepCoup, FA, Const1, Const2, Rate, RF, Floor, Cap, PF, PF1, PF2, PC, AB, NxtRst2, LookBackRate, LookBack, WamDate, PPUnits, PPCRShift, RcvCurrEscr, PayCurrEscr, 
	AggBookPrice, AmortMethod, ClientName, [Active?], Agency, [Int Coup Mult], [PP Coup Mult], [Sum Floor], [Sum Cap], [Servicing Model], [Loss Model], [Sched Cap?], Delay, 
	Ballon, IF, Mult1, Mult2, PrepaymentInterestRate )
SELECT s.AccountClass, s.AccountType, s.ScheduleType, s.PassThruRate, Last(s.Portfolio) AS Portfolio, CUSIP, Count(s.LoanNumber) AS H2, 
	Sum(s.Notional) AS AggNotional, Last(s.Mult) AS Mult, 
	IIf(Sum([CurrentLoanBalance])=0,Sum([OriginalAmount])/Sum([OriginalAmount]),Sum([CurrentLoanBalance])/Sum([OriginalAmount])) AS AggFactor, 
	Last(s.[Add Accrued?]) AS [Add Accrued?], Last(s.PV) AS PV, Last(s.Swap) AS Swap, 
	(Sum([Wac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWac, 
	(Sum([Coup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggCoup, 
	CInt(Sum([Wam]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWam, 
	(Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggAge, 
	IIf(Val([AccountType])=30 And [AggWam]<340,"Seasoned",IIf(Val([AccountType])=20 And [AggWam]<220,"Seasoned",IIf(Val([AccountType])=15 And [AggWam]<160,"Seasoned","MPF"))) AS H1, 
	Last(s.Swam) AS Swam, CInt(Sum([OWAM]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggOWAM, 
	"OAS" AS [P/O], 0 AS AggPrice, 0 AS OAS, Last(s.Settle) AS Settle, Last(s.Repo) AS Repo, Last(s.NxtPmt) AS NxtPmt, "AFT" AS PP, Last(s.PPConst) AS PPConst, 
	Last(s.PPMult) AS PPMult, Last(s.PPSense) AS PPSense, Last(s.Lag) AS Lag, Last(s.Lock) AS Lock, Last(s.PPFq) AS PPFq, Last(s.NxtPP) AS NxtPP, Last(s.NxtRst1) AS NxtRst1, 
	(Sum([PrepWac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepWac, 
	(Sum([PrepCoup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepCoup, 
	Last(s.FA) AS FA, Last(s.Const1) AS Const1, Last(s.Const2) AS Const2, Last(s.Rate) AS Rate, Last(s.RF) AS RF, Last(s.Floor) AS Floor, Last(s.Cap) AS Cap, Last(s.PF) AS PF, 
	Last(s.PF1) AS PF1, Last(s.PF2) AS PF2, Last(s.PC) AS PC, Last(s.AB) AS AB, Last(s.NxtRst2) AS NxtRst2, Last(s.LookBackRate) AS LookBackRate1, Last(s.LookBack2) AS LookBack2, 
	Last(CStr(s.[WamDate])) AS WamDate, Last(s.PPUnits) AS PPUnits, Avg(s.PPCRShift) AS PPCRShift, Last(s.RcvCurrEscr) AS RcvCurrEscr, Last(s.PayCurrEscr) AS PayCurrEscr, 
	Sum([CurrentLoanBalance]*[ChicagoParticipation]) AS AggBookPrice, 
	Last(s.AmortMethod) AS AmortMethod, Last(s.ClientName) AS ClientName, Last(s.[Active?]) AS [Active?], Last(s.Agency) AS Agency, Last(s.[Int Coup Mult]) AS [Int Coup Mult], 
	Last(s.[PP Coup Mult]) AS [PP Coup Mult], Last(s.[Sum Floor]) AS [Sum Floor], Last(s.[Sum Cap]) AS [Sum Cap], Last(s.[Servicing Model]) AS [Servicing Model], 
	Last(s.[Loss Model]) AS [Loss Model], Last(s.[Sched Cap?]) AS [Sched Cap?], Last(s.Delay) AS Delay, Last(s.Ballon) AS Ballon, Last(s.IF) AS IF, Last(s.Mult1) AS Mult1, 
	Last(s.Mult2) AS Mult2, (Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepaymentInterestRate
FROM qryUnhedgedPalmsDataSource as s
GROUP BY s.AccountClass, s.AccountType, s.ScheduleType, s.PassThruRate, s.OriginationYear, CUSIP
HAVING ((((Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])))<=30))
ORDER BY s.OriginationYear, s.AccountType DESC , s.ScheduleType DESC , (Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation]));
-- =================================================================== Aggregate MPF loans ======== END


	
-- qryAppendHedgedDataTotblSecuritiesPalmsSource (not used)
INSERT INTO tblSecuritiesPalmsSource ( AccountClass, AccountType, ScheduleType, PassThruRate, Portfolio, CUSIP, H2, AggNotional, Mult, AggFactor, [Add Accrued?], PV, Swap, AggWac, AggCoup, AggWam, AggAge, H1, Swam, AggOWAM, [P/O], AggPrice, OAS, Settle, Repo, NxtPmt, PP, PPConst, PPMult, PPSense, Lag, Lock, PPFq, NxtPP, NxtRst1, AggPrepWac, AggPrepCoup, FA, Const1, Const2, Rate, RF, Floor, Cap, PF, PF1, PF2, PC, AB, NxtRst2, LookBackRate, LookBack, WamDate, PPUnits, PPCRShift, RcvCurrEscr, PayCurrEscr, AggBookPrice, AmortMethod, ClientName, [Active?], Agency, [Int Coup Mult], [PP Coup Mult], [Sum Floor], [Sum Cap], [Servicing Model], [Loss Model], [Sched Cap?], Delay, Ballon, IF, Mult1, Mult2, PrepaymentInterestRate )
SELECT qryHedgedPalmsDataSource.AccountClass, qryHedgedPalmsDataSource.AccountType, qryHedgedPalmsDataSource.ScheduleType, qryHedgedPalmsDataSource.PassThruRate, Last(qryHedgedPalmsDataSource.Portfolio) AS Portfolio, Last(qryHedgedPalmsDataSource.CusipHedged) AS CUSIP, Count(qryHedgedPalmsDataSource.LoanNumber) AS H2, Sum(qryHedgedPalmsDataSource.Notional) AS AggNotional, Last(qryHedgedPalmsDataSource.Mult) AS Mult, IIf(Sum([CurrentLoanBalance])=0,Sum([OriginalAmount])/Sum([OriginalAmount]),Sum([CurrentLoanBalance])/Sum([OriginalAmount])) AS AggFactor, Last(qryHedgedPalmsDataSource.[Add Accrued?]) AS [Add Accrued?], Last(qryHedgedPalmsDataSource.PV) AS PV, Last(qryHedgedPalmsDataSource.Swap) AS Swap, (Sum([Wac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWac, (Sum([Coup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggCoup, CInt(Sum([Wam]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggWam, (Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggAge, IIf(Val([AccountType])=30 And [AggWam]<340,"Seasoned",IIf(Val([AccountType])=20 And [AggWam]<220,"Seasoned",IIf(Val([AccountType])=15 And [AggWam]<160,"Seasoned","MPF"))) AS H1, Last(qryHedgedPalmsDataSource.Swam) AS Swam, CInt(Sum([OWAM]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggOWAM, "OAS" AS [P/O], 0 AS AggPrice, 0 AS OAS, Last(qryHedgedPalmsDataSource.Settle) AS Settle, Last(qryHedgedPalmsDataSource.Repo) AS Repo, Last(qryHedgedPalmsDataSource.NxtPmt) AS NxtPmt, "AFT" AS PP, Last(qryHedgedPalmsDataSource.PPConst) AS PPConst, Last(qryHedgedPalmsDataSource.PPMult) AS PPMult, Last(qryHedgedPalmsDataSource.PPSense) AS PPSense, Last(qryHedgedPalmsDataSource.Lag) AS Lag, Last(qryHedgedPalmsDataSource.Lock) AS Lock, Last(qryHedgedPalmsDataSource.PPFq) AS PPFq, Last(qryHedgedPalmsDataSource.NxtPP) AS NxtPP, Last(qryHedgedPalmsDataSource.NxtRst1) AS NxtRst1, (Sum([PrepWac]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepWac, (Sum([PrepCoup]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepCoup, Last(qryHedgedPalmsDataSource.FA) AS FA, Last(qryHedgedPalmsDataSource.Const1) AS Const1, Last(qryHedgedPalmsDataSource.Const2) AS Const2, Last(qryHedgedPalmsDataSource.Rate) AS Rate, Last(qryHedgedPalmsDataSource.RF) AS RF, Last(qryHedgedPalmsDataSource.Floor) AS Floor, Last(qryHedgedPalmsDataSource.Cap) AS Cap, Last(qryHedgedPalmsDataSource.PF) AS PF, Last(qryHedgedPalmsDataSource.PF1) AS PF1, Last(qryHedgedPalmsDataSource.PF2) AS PF2, Last(qryHedgedPalmsDataSource.PC) AS PC, Last(qryHedgedPalmsDataSource.AB) AS AB, Last(qryHedgedPalmsDataSource.NxtRst2) AS NxtRst2, Last(qryHedgedPalmsDataSource.LookBackRate) AS LookBackRate1, Last(qryHedgedPalmsDataSource.LookBack2) AS LookBack2, Last(CStr([qryHedgedPalmsDataSource].[WamDate])) AS WamDate, Last(qryHedgedPalmsDataSource.PPUnits) AS PPUnits, Avg(qryHedgedPalmsDataSource.PPCRShift) AS PPCRShift, Last(qryHedgedPalmsDataSource.RcvCurrEscr) AS RcvCurrEscr, Last(qryHedgedPalmsDataSource.PayCurrEscr) AS PayCurrEscr, Sum([CurrentLoanBalance]*[ChicagoParticipation]) AS AggBookPrice, Last(qryHedgedPalmsDataSource.AmortMethod) AS AmortMethod, Last(qryHedgedPalmsDataSource.ClientName) AS ClientName, Last(qryHedgedPalmsDataSource.[Active?]) AS [Active?], Last(qryHedgedPalmsDataSource.Agency) AS Agency, Last(qryHedgedPalmsDataSource.[Int Coup Mult]) AS [Int Coup Mult], Last(qryHedgedPalmsDataSource.[PP Coup Mult]) AS [PP Coup Mult], Last(qryHedgedPalmsDataSource.[Sum Floor]) AS [Sum Floor], Last(qryHedgedPalmsDataSource.[Sum Cap]) AS [Sum Cap], Last(qryHedgedPalmsDataSource.[Servicing Model]) AS [Servicing Model], Last(qryHedgedPalmsDataSource.[Loss Model]) AS [Loss Model], Last(qryHedgedPalmsDataSource.[Sched Cap?]) AS [Sched Cap?], Last(qryHedgedPalmsDataSource.Delay) AS Delay, Last(qryHedgedPalmsDataSource.Ballon) AS Ballon, Last(qryHedgedPalmsDataSource.IF) AS IF, Last(qryHedgedPalmsDataSource.Mult1) AS Mult1, Last(qryHedgedPalmsDataSource.Mult2) AS Mult2, (Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])) AS AggPrepaymentInterestRate
FROM qryHedgedPalmsDataSource
GROUP BY qryHedgedPalmsDataSource.AccountClass, qryHedgedPalmsDataSource.AccountType, qryHedgedPalmsDataSource.ScheduleType, qryHedgedPalmsDataSource.PassThruRate, qryHedgedPalmsDataSource.OriginationYear, qryHedgedPalmsDataSource.PurchaseGroup, qryHedgedPalmsDataSource.CusipHedged
HAVING ((((Sum([Age]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation])))>30))
ORDER BY qryHedgedPalmsDataSource.OriginationYear, qryHedgedPalmsDataSource.AccountType DESC , qryHedgedPalmsDataSource.ScheduleType DESC , (Sum([PrepaymentInterestRate]*[AcctBalance]*[ChicagoParticipation])/Sum([AcctBalance]*[ChicagoParticipation]));

-- qryAppendHedgedDataTotblSecuritiesPalmsSource – NONSEASONED (not used)
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


-- =================================================================== Data for PolyPaths ======== BEGIN
-- DlyMPF ========== 
-- Hua 20150303 changed subActII and Account
-- Hua 20160211 use accountType instead of productCode to test if it's GNMBS (jumbo), SubActII
SELECT "DlyMPF" AS Query, "Asset MPF " & SubActII AS Account, "MBS" AS [Sec Type], "Asset" AS BSAccount, "MPF" AS SubActI, 
	IIf(Left([CUSIP],1) IN ("C", "F"), Left(ProductCode,2) & "2" & AccountType, 
		IIF(AccountType IN ("03","05","43","45"), IIf(RIGHT(AccountType,1)="3","GN30","GN15"), ProductCode)) AS SubActII, 
	CUSIP AS [CUST ID], tblSecuritiesPalmsSource.CUSIP AS CUSIP, AggOWAM AS [Orig Term], 
	[AggBookPrice]-IIf([month 1] Is Null,0,[Month 1]) AS Holding, AggCoup AS Coupon, 
	AggWac AS [WAC(coll)], [AggWam]-IIf([WAM_adj] Is Null,0,[WAM_Adj]) AS [WAM(coll)], 
	AggAge AS [WALA(coll)], Delay, 1 AS Factor, 
	AggPrice AS Price, Agency, "T+0" AS [Settlement Type], 
	"N" AS [Use Static Model], "30/360" AS DayCount, "USER" AS Source, "Price" AS PriceAnchor, 
	"tblSecuritiesPalmsSource in N:\Palms\BankDB\MPFFIVES for MiddlewareSQLandNewRemit.mdb" AS PriceSource, 
	qryWALoanSize_MPF.WALoanSize AS WALoanSize, qryWALoanSize_MPF.LTV
FROM (tblSecuritiesPalmsSource LEFT JOIN MRA_MPFPayDown ON tblSecuritiesPalmsSource.CUSIP=MRA_MPFPayDown.CUSTOM2) 
	LEFT JOIN qryWALoanSize_MPF ON tblSecuritiesPalmsSource.CUSIP=qryWALoanSize_MPF.CUSIP;
	
-- qryWALoanSize_MPF
-- Hua 20150303 Added ProductCode
SELECT CUSIP, Round(Sum(OriginalAmount/1000*cBal)/Sum(cBal),3) AS WALoanSize, Round(Sum(cBal),0) AS CurBal, 
	Round(100*Sum(OrigLTV*cBal)/Sum(cBal),5) AS LTV, AsOfDate, ProductCode
FROM qryMPFLoansAggregateSource
GROUP BY CUSIP, AsOfDate, ProductCode;

-- qryMPFLoansAggregateSource	
-- Hua 20150302 added prefix F and C to Cusip for v2 loans. Added ProductCode
-- Hua 20151223 updated the PassThruRate
-- Hua 20160213 updated PassThruRate for GNMBS, ScheduleType, CUSIP (drop the join on Hedges)
SELECT MAS.LoanNumber, 
	IIf(Left(ProductCode,2)="GL","GL",IIf(RemittanceTypeID=1,"SS",IIf(RemittanceTypeID=3,"MA","AA"))) AS ScheduleType, 
	IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]),Year([LoanRecordCreationDate])) AS OriginationYear, 
	Right([ProductCode],2) AS AccountType, 
	IIF(AccountType IN ("03","05","43","45"), INT((mas.InterestRate*200)+0.5)/2, CInt(mas.InterestRate*200)/2) AS PassThruRate, 
	OriginationYear AS AccountClass, 
	IIf(LoanInvestmentStatus="11","C", IIf(LoanInvestmentStatus="12","F","")) & ScheduleType & OriginationYear & AccountType & (PassThruRate*100) AS CUSIP, 
	OriginalAmount, ChicagoParticipation*CurrentLoanBalance AS CBal, Age, OrigLTV, EntryDate AS AsOfDate, ProductCode
FROM tblMPFLoansAggregateSource AS MAS

-- _DlyMPFDC ========== 
-- Hua 20150305 changed subActII to GN, GL, FX
-- Hua 20150611 Append "2" to account for MPFv2
-- Hua 20150911 use the Delay field from tblForwardCommitmentPalmsSource
-- Hua 20160212 updated SubActII, WALoanSize for GNMBS jumbo (=465 is the historical average of GN jumbo loans)
SELECT "DlyMPFDC" AS Query, "Asset MPFDC " & SubActII & IIf(left(CUSIP,1)="C","2","") AS Account, "MBS" AS [Sec Type], "Asset" AS BSAccount, "MPFDC" AS SubActI, 
	IIF(Right(ProductCode,2) IN ("03","05","43","45"),"GN",Left(ProductCode,2)) AS SubActII, 
	CUSIP AS [CUST ID], CUSIP, Owam AS [Orig Term], 
	AggNotional AS Holding, Coup AS Coupon, Wac AS [WAC(coll)], 
	Wam AS [WAM(coll)], Age AS [WALA(coll)], 1 AS Factor, P AS Price, 
	Agency, "N" AS [Use Static Model], "USER" AS [Settlement Type], Settle AS [Settle Date], "30/360" AS DayCount, 
	Delay, "USER" AS Source, "Price" AS PriceAnchor, 
	"tblForwardCommitmentPalmsSource in N:\Palms\BankDB\MPFFIVES for MiddlewareSQLandNewRemit.mdb" AS PriceSource, 
	IIf(productCode like "GL4?",465,IIf(ProductCode="FX30",179,IIf(productCode="FX15",132,IIf(productCode="FX20",143,IIf(productCode="GL30",144,130))))) AS WALoanSize
FROM tblForwardCommitmentPalmsSource;


-- DlyMPFDC
SELECT Query, Account, [Sec Type], BSAccount, SubActI, SubActII, [CUST ID], CUSIP, [Orig Term], Holding, Coupon, [WAC(coll)], [WAM(coll)], [WALA(coll)], 
	Factor, Price, Agency, [Use Static Model], [Settlement Type], [Settle Date], DayCount, Delay, Source, PriceAnchor, PriceSource, Round([_DlyMPFDC].WALoanSize,3) AS WALoanSize
FROM _DlyMPFDC;

-- polypaths Input ======= create the pf file by priceBatch
-- Hua 20150708 added the [factor date] for DCs
select [Sec Type],[BSAccount],[SubActI],[CUSIP],"" as [OCUSIP],[SubActII],"" as [Dated Date],"" as [Maturity],[Coupon],[Holding],[DayCount],"" as [Cpn~ Freq~],"" as [SwapCusip],
	"" as [Hedged],"" as [OAS],"" as [Call Date],"" as [Call Price],"" as [Put Date],"" as [Put Price],[Settlement Type],[Source],[Orig Term],[WAC(coll)],[WAM(coll)],[WALA(coll)],
	[Delay],[Factor],[Agency],[Use Static Model],[Price],"" as [Index],"" as [R Mult],"" as [Margin],"" as [First Coupon Date],"" as [Reset Freq~],"" as [DayCount (Pay)],
	"" as [Index(Pay)],"" as [P Mult],"" as [Coupon (Pay)],"" as [First Coupon Date (Pay)],"" as [Cpn~ Freq~ (Pay)],"" as [margin (Pay)],"" as [Reset Freq~ (Pay)],
	"" as [Settle Date],"" as [Strike],"" as [Swaption Swap Effective Date],"" as [Swaption Swap First Pay Date],"" as [Swaption Swap Termination Date],"" as [1st Exer~],
	"" as [Option Type],"" as [Option Exercise Type],"" as [Swaption Strike Rate],"" as [Option Exercise Notice],"" as [jrnlcode],"" as [Opening Market Value],[Cust ID],
	"" as [Sub Type],"" as [DTM],[Account],"" as [Cap],"" as [Floor],[PriceAnchor],[PriceSource],waLoanSize as [Avg Loan Size],[LTV],  "" as [Factor Date]
from DlyMPF
UNION 
select [Sec Type],[BSAccount],[SubActI],[CUSIP],"" as [OCUSIP],[SubActII],"" as [Dated Date],"" as [Maturity],[Coupon],[Holding],[DayCount],"" as [Cpn~ Freq~],"" as [SwapCusip],
	"" as [Hedged],"" as [OAS],"" as [Call Date],"" as [Call Price],"" as [Put Date],"" as [Put Price],[Settlement Type],[Source],[Orig Term],[WAC(coll)],[WAM(coll)],[WALA(coll)],
	[Delay],[Factor],[Agency],[Use Static Model],[Price],"" as [Index],"" as [R Mult],"" as [Margin],"" as [First Coupon Date],"" as [Reset Freq~],"" as [DayCount (Pay)],
	"" as [Index(Pay)],"" as [P Mult],"" as [Coupon (Pay)],"" as [First Coupon Date (Pay)],"" as [Cpn~ Freq~ (Pay)],"" as [margin (Pay)],"" as [Reset Freq~ (Pay)],
	[Settle Date],"" as [Strike],"" as [Swaption Swap Effective Date],"" as [Swaption Swap First Pay Date],"" as [Swaption Swap Termination Date],"" as [1st Exer~],
	"" as [Option Type],"" as [Option Exercise Type],"" as [Swaption Strike Rate],"" as [Option Exercise Notice],"" as [jrnlcode],"" as [Opening Market Value],[Cust ID],
	"" as [Sub Type],"" as [DTM],[Account],"" as [Cap],"" as [Floor],[PriceAnchor],[PriceSource],waLoanSize as [Avg Loan Size],"" as [LTV], 
	DateSerial(Year(Date()), Month(Date()), 1) as [Factor Date]
from DlyMPFDC;
-- =================================================================== Data for PolyPaths ======== END


--------------- 20140820
-- ================================================= cmdInsertMiddleware_Click  ======== begin
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
EXEC usp_DataLoadStatistics_INSERT "Daily MPF Loans"
-- ================================================= cmdInsertMiddleware_Click  ======== END




-- FHLBNav_Admin
SELECT FHLBNAV_User.username
FROM FHLBNAV_User
WHERE (((FHLBNAV_User.Admin)=Yes));

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

-- qryUpdateProductCodeforFX15bookedthirty --------- No use
UPDATE tblMPFLoansAggregateSource SET tblMPFLoansAggregateSource.ProductCode = "FX15"
WHERE (((tblMPFLoansAggregateSource.ProductCode)="FX30") AND ((tblMPFLoansAggregateSource.NumberOfMonths)<=241 And (tblMPFLoansAggregateSource.NumberOfMonths)>=239));


-- =============================Run PALMS Data Report
-- qryUpdateFlexiSwap -- not used any more -- Hua 20150309
-- UPDATE tblFlexiSwap, tblMPFLoansAggregateSource SET tblMPFLoansAggregateSource.ProductCode = "MS" & Right([ProductCode],2);

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

-- UpdateMPF15&20PPMult
UPDATE tblSecuritiesPalmsSource SET tblSecuritiesPalmsSource.PPMult = 1.3
WHERE (((tblSecuritiesPalmsSource.AccountType)="15" Or (tblSecuritiesPalmsSource.AccountType)="20") AND ((tblSecuritiesPalmsSource.AggWac)>=6.02));

-- UpdateMPF30PPMult
UPDATE tblSecuritiesPalmsSource SET tblSecuritiesPalmsSource.PPMult = 1.2
WHERE (((tblSecuritiesPalmsSource.AccountType)="30") AND ((tblSecuritiesPalmsSource.AggWac)>=7.04));

-- UpdateMPF15&20PPMultMarkToModel
UPDATE tblMarkToModelLoans SET tblMarkToModelLoans.PPMult = 1.3
WHERE (((tblMarkToModelLoans.AccountType)="15" Or (tblMarkToModelLoans.AccountType)="20") AND ((tblMarkToModelLoans.AggWac)>=6.02));

-- UpdateMPF30PPMultMarkToModel
UPDATE tblMarkToModelLoans SET tblMarkToModelLoans.PPMult = 1.2
WHERE (((tblMarkToModelLoans.AccountType)="30") AND ((tblMarkToModelLoans.AggWac)>=7.04));

---- ========== Export data to Palm Production (not used)


-- ============= clear tables before compact === Command76_Click()
-- qryDeletetblLoanFundingDateViewtoTABLE
DELETE tblLoanFundingDateViewTABLE.*
FROM tblLoanFundingDateViewTABLE;

-- qryDeletetblMPFLoansAggregateSource
DELETE tblMPFLoansAggregateSource.*
FROM tblMPFLoansAggregateSource;

-- qryDeletetbltempMPF
DELETE tblTempMPF.*
FROM tblTempMPF;
-- ============= clear tables before compact == End


-- qryMissingHFSPrices ======= cmdCheckMissingHFSPrices_Click
SELECT Cusip, LoanNumber, Price
FROM tblHFSLoansEOD
WHERE (Price Is Null) OR (Price = 0);

-- qryDeleteMPFandForwardsfromInstrumenttblMiddleware ===== cmdDeleteMiddleware_Click
--DELETE FROM Instrument_MBS_MPF
--WHERE (((Instrument_MBS_MPF.CUSIP) Like "AA%" Or (Instrument_MBS_MPF.CUSIP) Like "GL%" Or (Instrument_MBS_MPF.CUSIP) Like "MS%" Or (Instrument_MBS_MPF.CUSIP) Like "SS%" Or (Instrument_MBS_MPF.CUSIP) Like "DC%" Or (Instrument_MBS_MPF.CUSIP) Like "HC%" Or (Instrument_MBS_MPF.CUSIP) Like "MA%") AND ((Instrument_MBS_MPF.BondCUSIP)="MPF" Or (Instrument_MBS_MPF.BondCUSIP)="MPFForward" Or (Instrument_MBS_MPF.BondCUSIP)="MPFAct"))
--

-- qryDeleteMPFandForwardsfromInstrumenttblMiddleware ===== cmdDeleteMiddleware_Click
-- Hua 20150310 dropped the cusip filters
DELETE FROM Instrument_MBS_MPF
WHERE (BondCUSIP="MPF" Or BondCUSIP="MPFForward" Or BondCUSIP="MPFAct")


-- qryDeleteDataFromPortfolioContentsMiddleware ========= cmdDeleteMiddleware_Click()
DELETE FROM Portfolio_Contents_MBS_MPF

-- qryHFSLoanPricesToTPODS ========= cmdExportHFSToTPODS_Click
SELECT FORMAT(Date(),"yyyymmdd") AS PricingDate, tblHFSLoansEOD.LoanNumber, tblHFSLoansEOD.Price
FROM tblHFSLoansEOD;

-- qryUpdatetblMonthEndDate
UPDATE tblMonthEndDate SET tblMonthEndDate.MonthEndDate = [Forms]![frmSetEOMSchedule]![txtEndofMonthScheduleDate];


----=============================== pricing 201409
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
WHERE ((([ScheduleType]+[AccountClass]+Left([AccountType],1)+"0"+CStr(([PassThruRate]*100)))=[MPFPrice].[CUSIP]));

-- ========== Collect MPF forward level data
-- qryUpdateMPFForwardPrices --- Hua 20150604 droped the [schedule type] filter 
UPDATE ForwardSettleDCs as f, qryMPFForwardPrice as p SET f.Price = p.[Price]
WHERE (((p.[Date])=CDate(f.[DeliveryDate])) AND ((CStr(f.[NoteRate]))=p.[Rate]) 
AND (f.[RemittanceTypeID]=p.[ServicingRemittanceType])
AND ((f.ProductCode)=p.[CatType]))

-- qryUpdateBackwordPriceTo100 ------------- called in the form: frmShowMissingMPFForwardPrice
UPDATE ForwardSettleDCs SET ForwardSettleDCs.Price = 100
WHERE (((ForwardSettleDCs.Price)=0) AND ((ForwardSettleDCs.DeliveryDate)<Now()));

-- qryMPFForwardPrice 
SELECT CatType, ProdType, Year, Date, [MBS Coupon], IIf([MPF Rate] Is Null,Null,CStr([MPF Rate]/100)) AS Rate, Day, Price, ServicingRemittanceType
FROM MPFForwardPrice
WHERE ((CatType) Is Not Null) AND ((ProdType) Is Not Null) AND ((Year) Is Not Null) AND ((Date) Is Not Null) AND (([MBS Coupon]) Is Not Null) 
AND ((IIf([MPF Rate] Is Null,Null,CStr([MPF Rate]/100))) Is Not Null) AND ((ServicingRemittanceType) Is Not Null) AND ((Price) Is Not Null);
----=============================== pricing 201409


-- qryDeleteDataFromTblForwardCommitmentPalmsSource
DELETE tblForwardCommitmentPalmsSource.*
FROM tblForwardCommitmentPalmsSource;

-- qryDeleteDataFromTblForwardCommitmentMarkToModelLoans
DELETE tblForwardCommitmentMarkToModel.*
FROM tblForwardCommitmentMarkToModel;








---------- scratch:
-- 20121220 qryLoanCusipNew
-- Hua 20150303 Added prefix C/F
-- Hua 20160229 updated PassThruRate (cusip), Agency to handle GNMBS jumbo loans; deleted owam
SELECT LoanNumber, IIf(Left([ProductCode],2)="MS","MS",IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=3,"MA","AA")))) AS ScheduleType, 
	IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]),Year([LoanRecordCreationDate])) AS OriginationYear, Right([ProductCode],2) AS AccountType, 
	IIF(AccountType IN ("03","05","43","45"), INT((InterestRate*200)+0.5)/2, CInt(InterestRate*200)/2) AS PassThruRate, 
	IIf(LoanInvestmentStatus="11","C",
		IIf(LoanInvestmentStatus="12","F","")) & [ScheduleType] & [OriginationYear] & [AccountType] & ([PassThruRate]*100) AS Cusip, 
	[MPFBalance]*[ChicagoParticipation] AS Notional, (CInt([InterestRate]*100000)/100000)*100 AS Wac, (CInt([Coupon]*100000)/100000)*100 AS Coup, 
	CInt(NPer(([InterestRate]/12),([PIAmount]*-1),[CurrentLoanBalance])) AS WAM, Age, 
	EntryDate AS Settle, CurrentLoanBalance*ChicagoParticipation AS cBal, 
	IIf([ScheduleType]="GL", IIf(AccountType IN ("03","05","43","45"),"GNMA2","GNMA"),IIf([ScheduleType]="MS","GNMA","FNMA")) AS Agency,
	OriginalAmount, ChicagoParticipation, CEFee, DeliveryCommitmentNumber, PIAmount, interestRate, maNumber, PFINumber, ProductCode, origLTV
FROM tblMPFLoansAggregateSource AS l;



-- remittance type adjustment
=IF($A74="","",IF(LEFT($D74,2)="AA","S/R",IF(LEFT($D74,2)="MA","A/A","S/S")))
=IF($A74="","",IF($H74="S/S",0,IF($H74="S/R",-0.33,0.02)))

-- qryLoanSize (Hua)
SELECT Cusip, Sum(OriginalAmount/1000*cBal)/Sum(cBal) AS WAloanSize, settle, Sum(cBal) AS curBal
FROM qryLoanCusip
GROUP BY Cusip, settle;

-- qryFwdLoanSize (Hua)
SELECT CUSIP, IIF(ProductCode="FX30", 179, IIF(productCode="FX15", 132, IIF(productCode="FX20", 143, IIF(productCode="GL30", 144, 130)))) as WALoanSize
FROM tblForwardCommitmentPalmsSource;


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
* AI Iteration 2		01/29/2014		Avdhut Vaidya		TFS#4328 :: item#1658 :: Ai Phase 1 :: using an "iNNER JOiN" instead of "LEFT JOiN" while
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


----- [mtqTblNorwestLoans LASPymt 5mindelay] in norwest running on 18th each month
SELECT LoanNumber, PoolNumber, PFINumber, ScheduleEndPrincipalBal AS [Sched End Prin Bal], OriginalPrincipalBal, LoanAdminFee, CEFee, MPFBankCEFee, LiquidationAmount 
INTO tblNorwestLoans
FROM Dbo_UV_MaxPaydate
ORDER BY LoanNumber;



-- dbo.UV_MaxPayDate  in MPFFHLBDW
CREATE VIEW [dbo].[UV_MaxPayDate]
AS
SELECT     PH.LoanNumber, PH.PoolNumber, PH.PFINumber, PH.ScheduleEndPrincipalBal, PH.OriginalPrincipalBal, PH.LoanAdminFee, PH.CEFEE, 
                      PH.MPFBankCEFee, PH.LiquidationAmount
FROM         dbo.UV_LAS_PaymentsHistory AS PH INNER JOIN
                          (SELECT     LoanNumber, MAX(PayDate) AS MaxofPayDate
                            FROM          dbo.UV_LAS_PaymentsHistory
                            GROUP BY LoanNumber) AS MaxPayDate ON MaxPayDate.LoanNumber = PH.LoanNumber AND MaxPayDate.MaxofPayDate = PH.PayDate

							
							
/*
****** OBJECT: UV_DCParticipation_MRA_NOSHFD **************************************************************************************
* Description :	Object created for Bank's Downstream Views
* --------------------------------------------------------------------------------------------------------------------------
* Version           Date of Change         Author               Description
* ----------------- ---------------------- -------------------- ------------------------------------------------------------
* DM 11.0			01/28/2013				Kiran Nair			Item#1662 :: TFS# 4340 :: GNMA E2E_AI EUC Testing ::
																Created DC participation user view for MRA
****************************************************************************************************************************
*/
CREATE VIEW UV_DCParticipation_MRA_NOSHFD
AS
	SELECT	TOP	100	percent
			UV_Schedule_NOSHFD.ProductCode, 
			UV_QryDeliveryCommitment.NoteRate, 
			UV_LoanRatesAgentFees_NOSHFD.Fee, 
			UV_QryDeliveryCommitment.DeliveryStatus, 
			UV_Schedule_NOSHFD.ScheduleType, 
			UV_Schedule_NOSHFD.RemittanceTypeID,
			UV_QryDeliveryCommitment.EntryDate, 
			UV_QryDeliveryCommitment.EntryTime, 
			UV_QryDeliveryCommitment.DeliveryDate, 
			UV_PFI_NOSHFD.Name AS FullName, 
			UV_QryDeliveryCommitment.PFINumber, 
			UV_QryDeliveryCommitment.MANumber, 
			UV_QryDeliveryCommitment.DeliveryCommitmentNumber, 
			UV_QryDeliveryCommitment.DeliveryAmount, 
			UV_QryDeliveryCommitment.FundedAmount,
			UV_QryDeliveryCommitment.ParticipationOrgKey,
			UV_QryDeliveryCommitment.ParticipationOrgName,
			CAST(UV_QryDeliveryCommitment.DeliveryAmount * UV_QryDeliveryCommitment.Participation AS money) AS DeliveryAmountP,
			CAST(UV_QryDeliveryCommitment.FundedAmount * UV_QryDeliveryCommitment.Participation AS money) AS FundedAmountP, 
			UV_QryDeliveryCommitment.LastUpdatedDate, 
			UV_Schedule_NOSHFD.ScheduleCode, 
			UV_QryDeliveryCommitment.Participation,
			UV_QryDeliveryCommitment.IsExtended, 
			UV_QryDeliveryCommitment.ServicingFee, 
			UV_QryDeliveryCommitment.ExcessServicingFee, 
			UV_MasterCommitment_NOSHFD.CEFee, 
			UV_MasterCommitment_NOSHFD.CEPerformanceFee
	
	FROM	UV_QryDeliveryCommitment 
			INNER JOIN LOS_Prod.dbo.Schedule UV_Schedule_NOSHFD 
				ON UV_QryDeliveryCommitment.ScheduleCode = UV_Schedule_NOSHFD.ScheduleCode
			INNER JOIN LOS_Prod.dbo.Org UV_PFI_NOSHFD 
				ON CAST(UV_QryDeliveryCommitment.PFINumber AS VARCHAR(6)) = UV_PFI_NOSHFD.OrgID
			INNER JOIN  LOS_Prod.dbo.MasterCommitment UV_MasterCommitment_NOSHFD
				ON UV_QryDeliveryCommitment.MANumber = UV_MasterCommitment_NOSHFD.MasterCommitmentID
			INNER JOIN LOS_Prod.dbo.Product UV_Product_NOSHFD 
				ON UV_Schedule_NOSHFD.ProductCode = UV_Product_NOSHFD.ProductCode
			LEFT JOIN LOS_Prod.dbo.LoanRatesAgentFees UV_LoanRatesAgentFees_NOSHFD 
				ON UV_QryDeliveryCommitment.NoteRate = UV_LoanRatesAgentFees_NOSHFD.Rate
				AND UV_QryDeliveryCommitment.ScheduleCode = UV_LoanRatesAgentFees_NOSHFD.ScheduleCode

	ORDER	BY	UV_QryDeliveryCommitment.NoteRate,
				UV_QryDeliveryCommitment.DeliveryDate
  
 
							