
Moving to PRODUCTION:
1) replace "CInt(NPer(([a.InterestRate]/12),([a.PIAmount]*-1),[a.CurrentLoanBalance])) as WAM" for NumberOfMonths-age as WAM in qryPalmsDataSource
2) Uncomenting FileName/TPODSLoc and fso.CopyFile in cmdExportHFSToTPODS_Click()
3) Change all the links from _T6 to MPFFHLBDW
4) Replace "and LOANLoanInvestmentStatus <> "09" " for ???? in qrySelectAllLoans
5) In MPF-pricing-template.MPF-pricing 
CUSIP FINAL (D)
=IF($A496="","",IF(OR(LEFT($A496,1)="C",LEFT($A496,1)="F"),RIGHT($A496,11),IF(LEN($A496)=14,CONCATENATE(LEFT($A496,6),MID($A496,9,5)),$A496)))
RawCusip (E) (?????)
=IF($A496="","",CONCATENATE(IF(LEFT($D496,2)<>"GL",IF(LEFT($D496,2)<>"MS","FN", "GN"), "GN"),IF(MID($D496,7,2)="15","10",IF(MID($D496,7,2)="20","30",IF(MID($D496,7,2)="03","30",IF(MID($D496,7,2)="05","10","30")))),VALUE(MID($D496,9,2))-5,"S",MID($D496,3,4)))
6) the link of MPFPrice



--==========  201508
-------- Hua 20150810 assume dbo_UV_FairValueOptionLoans_NOSHFD contains real time data, and "06" records included
	SELECT 	f.LoanNumber,
		f.LoanInvestmentStatus,
		f.ChicagoParticipation,
		f.InterestRate,
		f.FundingDate,
		f.ProductCode,
		f.Action_Code,
		IIf(Left(f.ProductCode,2)="GL","GL",
			IIf(f.RemittanceTypeID=1,"SS",
			IIf(f.RemittanceTypeID=2,"AA", "MA"))) AS ScheduleType, 
		INT((f.InterestRate*200)+0.5)/2 AS PassThruRate,
		Right(f.ProductCode,2) AS AccountType,
		NZ(Year(mpf.LoanRecordCreationDate),Year(Date())) AS OriginationYear,
		IIf(f.LoanInvestmentStatus="12","F","") & (ScheduleType & OriginationYear & AccountType & PassThruRate*100) AS CUSIP,
		Cdbl(0) AS Price		
	INTO tblHFSLoansEOD
	FROM dbo_UV_FairValueOptionLoans_NOSHFD AS f LEFT JOIN tblMPFLoansAggregateSource AS mpf ON f.LoanNumber = mpf.LoanNumber		
	WHERE f.LoanInvestmentStatus NOT IN ("03", "11")
	AND   (f.LOANSettlementDate is NULL OR Date()-1<f.LOANSettlementDate )

-------- qryAppendV2HFVloans  
-- Hua 20150809 query not needed any more cause the union in qryInsertHFSLoansEOD
INSERT INTO tblHFSLoansEOD (loanNumber)
select LoanNumber 
from  [MPF Hedges]
where loanNumber<0


----- qryCheckExistingHFSLoanCount
-- Hua 20150809 change dbo_UV_HFSLoans_NOSHFD to dbo_UV_FairValueOptionLoans_NOSHFD
SELECT l.ProductCode, Count(l.LoanNumber) AS CountOfExitingLoans
FROM tblMPFLoansAggregateSource as l INNER JOIN dbo_UV_FairValueOptionLoans_NOSHFD as f ON l.LoanNumber = f.LoanNumber
GROUP BY l.ProductCode
HAVING (((l.ProductCode) In ("GL03","GL05")));


---------- qryPalmsDataSource:  
-- Hua 20150812 make age>0
SELECT "MPF" AS Portfolio, 
	IIf(Left([a.ProductCode],2)="MS","MS",IIf(Left([a.ProductCode],2)="GL","GL",IIf([a.RemittanceTypeID]=1,"SS",IIf([a.RemittanceTypeID]=3,"MA","AA")))) AS ScheduleType, 
	True AS [Active?], 
	IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]), Year([a.LoanRecordCreationDate])) AS OriginationYear, 
	a.LoanRecordCreationDate, a.LoanNumber, a.DeliveryCommitmentNumber, 
	IIf([PortfolioIndicator]="BATCH", IIf(Month([OriginalLoanClosingDate])>9,"","0") & Month([OriginalLoanClosingDate]),
		IIf(Month([a.LoanRecordCreationDate])>9,"","0") & Month([a.LoanRecordCreationDate])) AS OriginationMonth, 
	Right([a.ProductCode],2) AS AccountType, 
	IIF(AccountType = "03" OR AccountType = "05", INT((a.InterestRate*200)+0.5)/2, (CInt([a.InterestRate]*200)/2)) AS PassThruRate, 
	IIf(LoanInvestmentStatus="11","C",
		IIf(LoanInvestmentStatus="12","F","")) & ScheduleType & OriginationYear & AccountType & (PassThruRate*100) AS CUSIP, 
	[OriginationYear] AS AccountClass, [a.MPFBalance]*[a.ChicagoParticipation] AS Notional, 
	1 AS Mult, IIf([a.CurrentLoanBalance]=0,1,[a.CurrentLoanBalance]/[a.OriginalAmount]) AS Factor, 
	"1" AS [Add Accrued?], "Mid" AS PV, "Mid" AS Swap, (CInt([a.InterestRate]*100000)/100000)*100 AS Wac, (CInt([a.Coupon]*100000)/100000)*100 AS Coup, [Wac]-[Coup] AS Diff,
	
	IIf(a.age<2, numberofmonths-1, CInt(NPer(([a.InterestRate]/12),([a.PIAmount]*-1),[a.CurrentLoanBalance]))) AS WAM, 
	IIf(a.age<2, 1, a.age) as Age,
	
	IIf([AccountType]=15,160,IIf([AccountType]=20,220,IIf([AccountType]=30,335,IIf([AccountType]=03,335,IIf([AccountType]=05,160,""))))) AS Swam, 
	tblPALMSRepo.RepoRate AS Repo, 
	a.EntryDate AS Settle, NextDate([Settle],[ScheduleType]) AS NxtPmt, "0" AS PPConst, "1" AS PPMult, "1" AS PPSense, "0" AS Lag, "0" AS Lock, 
	"12" AS PPFq, [NxtPmt] AS NxtPP, NextDate([Settle],[ScheduleType]) AS NxtRst1, [Wac] AS PrepWac, [Coup] AS PrepCoup, "Fixed" AS FA, 
	"0" AS Const2, 1 AS Mult1, 0 AS Mult2, "3L" AS Rate, "12" AS RF, -1000000000 AS Floor, 1000000000 AS Cap, 1 AS PF, 1000000000 AS PF1, 
	1000000000 AS PF2, 1000000000 AS PC, "Bond" AS AB, [NxtRst1] AS NxtRst2, 0 AS LookBack2, 
	0 AS LookBackRate, EntryDate AS WamDate, "CPR" AS PPUnits, 0 AS PPCRShift, 0 AS RcvCurrEscr, 0 AS PayCurrEscr, "StraightLine" AS AmortMethod, 
	"MPFProgram" AS ClientName, a.PrepaymentInterestRate, a.CurrentLoanBalance AS AcctBalance, 
	IIf([ScheduleType]="GL", IIf(AccountType = "03" OR AccountType = "05","GNMA2","GNMA"),IIf([ScheduleType]="MS","GNMA","FNMA")) AS Agency,
	IIf([ScheduleType]="MS",18,IIf([ScheduleType]="GL",18,IIf([ScheduleType]="MA",2,IIf([ScheduleType]="AA",48, 18)))) AS Delay, 
	a.NumberOfMonths AS OWAM, 0 AS Ballon, 1 AS IF, 0 AS Const1, 0 AS [Int Coup Mult], 1 AS [PP Coup Mult], -10000000000 AS [Sum Floor], 
	10000000000 AS [Sum Cap], "None" AS [Servicing Model], "None" AS [Loss Model], 0 AS [Sched Cap?], a.CurrentLoanBalance, 
	a.OriginalAmount, a.ChicagoParticipation
FROM tblMPFLoansAggregateSource AS a, tblPALMSRepo


-- 20150911 GL loan RemittanceTypeID distribution (DC)
SELECT dbo_UV_DCParticipation_MRA_NOSHFD.ProductCode, dbo_UV_DCParticipation_MRA_NOSHFD.RemittanceTypeID, Sum(dbo_UV_DCParticipation_MRA_NOSHFD.DeliveryAmount) AS SumOfDeliveryAmount
FROM dbo_UV_DCParticipation_MRA_NOSHFD
GROUP BY dbo_UV_DCParticipation_MRA_NOSHFD.ProductCode, dbo_UV_DCParticipation_MRA_NOSHFD.RemittanceTypeID
HAVING (((dbo_UV_DCParticipation_MRA_NOSHFD.ProductCode) Like "GL*"));



