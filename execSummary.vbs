

'=============== 20110414 Executive Summary data processing
1) download the FMHPI file from FredieMac and Paste it into h:\MPF\prepayRpt\data\FMHPI.xls
2) In db execSummary.mdb, appFMHPI (After truncate the 	 in SQLSEVER)
3) Run the following query in SQLSERVER after truncating FMHPIstate:

INSERT INTO FMHPIstate 
SELECT * FROM (
select 'AK' as state, yrMon, AK as HPI from FMHPI UNION
select 'AL' as state, yrMon, AL as HPI from FMHPI UNION
select 'AR' as state, yrMon, AR as HPI from FMHPI UNION
select 'AZ' as state, yrMon, AZ as HPI from FMHPI UNION
select 'CA' as state, yrMon, CA as HPI from FMHPI UNION
select 'CO' as state, yrMon, CO as HPI from FMHPI UNION
select 'CT' as state, yrMon, CT as HPI from FMHPI UNION
select 'DC' as state, yrMon, DC as HPI from FMHPI UNION
select 'DE' as state, yrMon, DE as HPI from FMHPI UNION
select 'FL' as state, yrMon, FL as HPI from FMHPI UNION
select 'GA' as state, yrMon, GA as HPI from FMHPI UNION
select 'HI' as state, yrMon, HI as HPI from FMHPI UNION
select 'IA' as state, yrMon, IA as HPI from FMHPI UNION
select 'ID' as state, yrMon, ID as HPI from FMHPI UNION
select 'IL' as state, yrMon, IL as HPI from FMHPI UNION
select 'IN' as state, yrMon, [IN] as HPI from FMHPI UNION
select 'KS' as state, yrMon, KS as HPI from FMHPI UNION
select 'KY' as state, yrMon, KY as HPI from FMHPI UNION
select 'LA' as state, yrMon, LA as HPI from FMHPI UNION
select 'MA' as state, yrMon, MA as HPI from FMHPI UNION
select 'MD' as state, yrMon, MD as HPI from FMHPI UNION
select 'ME' as state, yrMon, ME as HPI from FMHPI UNION
select 'MI' as state, yrMon, MI as HPI from FMHPI UNION
select 'MN' as state, yrMon, MN as HPI from FMHPI UNION
select 'MO' as state, yrMon, MO as HPI from FMHPI UNION
select 'MS' as state, yrMon, MS as HPI from FMHPI UNION
select 'MT' as state, yrMon, MT as HPI from FMHPI UNION
select 'NC' as state, yrMon, NC as HPI from FMHPI UNION
select 'ND' as state, yrMon, ND as HPI from FMHPI UNION
select 'NE' as state, yrMon, NE as HPI from FMHPI UNION
select 'NH' as state, yrMon, NH as HPI from FMHPI UNION
select 'NJ' as state, yrMon, NJ as HPI from FMHPI UNION
select 'NM' as state, yrMon, NM as HPI from FMHPI UNION
select 'NV' as state, yrMon, NV as HPI from FMHPI UNION
select 'NY' as state, yrMon, NY as HPI from FMHPI UNION
select 'OH' as state, yrMon, OH as HPI from FMHPI UNION
select 'OK' as state, yrMon, OK as HPI from FMHPI UNION
select 'OR' as state, yrMon, [OR] as HPI from FMHPI UNION
select 'PA' as state, yrMon, PA as HPI from FMHPI UNION
select 'RI' as state, yrMon, RI as HPI from FMHPI UNION
select 'SC' as state, yrMon, SC as HPI from FMHPI UNION
select 'SD' as state, yrMon, SD as HPI from FMHPI UNION
select 'TN' as state, yrMon, TN as HPI from FMHPI UNION
select 'TX' as state, yrMon, TX as HPI from FMHPI UNION
select 'UT' as state, yrMon, UT as HPI from FMHPI UNION
select 'VA' as state, yrMon, VA as HPI from FMHPI UNION
select 'VT' as state, yrMon, VT as HPI from FMHPI UNION
select 'WA' as state, yrMon, WA as HPI from FMHPI UNION
select 'WI' as state, yrMon, WI as HPI from FMHPI UNION
select 'WV' as state, yrMon, WV as HPI from FMHPI UNION
select 'WY' as state, yrMon, WY as HPI from FMHPI UNION
select 'US' as state, yrMon, US as HPI from FMHPI ) AS A

4) insert the last survey rates (Freddie) of the Month into tblSurveyRate
-- 5) run mkTblOBA and Then qryProdOBA to Get openBasis
6) run mkTblLoanStrip (Last businessDate, HAUS2_archive) and Then qryProdStrip to Get agent fee and closed basis. Paste and Adjust source and Refresh the PT
7) run qry(Hist)ProdInfo (MPFREPOS) to Get bal, GNR, prepayInc, WALA, FICO, oLTV. Paste and transpo
8) run qryCurLTV to Get the curLTV. Paste and transpo
9) run qryIncBal to Get the balance of each incentive level. Paste in PT Tab and Refresh the PT


'-------- mkTblOBA
SELECT HLFLoanNbr, HLFFHLBCIncepPLAmt, HLFLoanPaydownAdjAmt, HLFEffectivenessInd, HLFHedgeID INTO tblOBA
FROM dbo_HAUS_Loan_Fact
WHERE ((HLFAsOfDate)=[asOfDate]) And ((HLFEffectivenessInd)="1" Or (HLFLoanPaydownAdjAmt<>0 And (HLFHedgeID Is Null)));

'--------- qryProdOBA
SELECT l.productCode, Sum(HLFFHLBCIncepPLAmt+HLFLoanPaydownAdjAmt)/1000000 AS OBA
FROM   tblOBA AS f INNER JOIN dbo_UV_DM_MPF_Daily_Loan AS l ON f.HLFLoanNbr = l.LoanNumber
WHERE  l.ChicagoParticipation>0	
and    l.Sched_End_Prin_Bal>0
GROUP BY l.productCode;

'----- mkTblLoanStrip
SELECT HAILoanNbr, HAIAmortCode, HAIUnamortizedNetAmt 
into   tblLoanStrip
FROM dbo_arch_Amortization_Information  
WHERE (HAIAmortCode)<>"D"
AND   (HAIasOfDate)=[asOfDate]


'--------- qryProdStrip   20140603
--- SELECT a.HAIAmortCode, i.productCode, Sum(a.HAIUnamortizedNetAmt)/1000000 AS netAmt
--- FROM   tblLoanStrip AS a INNER JOIN qryLnInfo AS i ON a.HAILoanNbr=i.LoanNumber
--- GROUP BY a.HAIAmortCode, i.productCode
--- ORDER BY i.productCode, a.HAIAmortCode;
SELECT a.HAIAmortCode, l.productCode, Sum(a.HAIUnamortizedNetAmt)/1000000 AS netAmt
FROM   tblLoanStrip AS a INNER JOIN tblMPFLoansAggregateSource20140618 AS l ON a.HAILoanNbr=l.LoanNumber
GROUP BY a.HAIAmortCode, l.productCode
ORDER BY l.productCode, a.HAIAmortCode;





'---- qryProdInfo
SELECT ProductCode, Sum(curBal)/1000000 AS chBal, Sum(Rate*curBal)/1000000/chBal AS GNR, Sum(incBps*curBal)/1000000/chBal AS incentive, Sum(age*curBal)/1000000/chBal AS WALA, 
   Sum(FICO*curBal)/1000000/chBal AS oFICO, Sum(LTV*curBal)/1000000/chBal AS oLTV, Sum(curLTV*curBal)/1000000/chBal AS cLTV
FROM qryLnInfo
GROUP BY ProductCode;

'-------- qryCurLTV 3 month delay
-- SELECT m.productCode, Sum(m.curBal*m.Sched_End_Prin_Bal*(m.LTV/m.LoanAmount)*(o.HPI/c.HPI))/Sum(m.curBal) AS curLTV
-- FROM   qryLnInfo AS m, dbo_FMHPIstate AS o, dbo_FMHPIstate AS c
-- WHERE  c.yrMon=(Select Max(yrMon) from dbo_FMHPIstate)   
--    AND    o.state=c.state
--    AND    m.PropertyState=o.state
--    AND    o.yrMon=Year(DateAdd("m", -4, FirstPaymentDate))&"M"&IIf(Month(DateAdd("m", -4, FirstPaymentDate))<10,"0","")&Month(DateAdd("m", -4, FirstPaymentDate))
-- GROUP BY m.productCode;

SELECT m.productCode, Sum(m.curBal*m.Sched_End_Prin_Bal*(m.LTV/m.LoanAmount)*(o.HPI/c.HPI))/Sum(m.curBal) AS curLTV
FROM tmpLnInfo AS m, dbo_FMHPIstate AS o, dbo_FMHPIstate AS c
WHERE c.yrMon="2014M06"  
   AND    o.state=c.state
   AND    m.PropertyState=o.state
   AND    o.yrMon=Year(DateAdd("m", -4, FirstPaymentDate))&"M"&IIf(Month(DateAdd("m", -4, FirstPaymentDate))<10,"0","")&Month(DateAdd("m", -4, FirstPaymentDate))
GROUP BY m.productCode;





'------- qryIncBal
-- SELECT ProductCode, incUpper, Sum(curBal)/1000000 AS cBal
-- FROM qryLnInfo
-- GROUP BY ProductCode, incUpper
-- ORDER BY ProductCode, incUpper;
SELECT  l.ProductCode,  
	Int(4*(100*[InterestRate]-IIf(NumberOfMonths<240,r.FN15,r.FN30)))*25+25 AS incUpper ,
	sum(curBal)/1000000 AS cBal
FROM  tmpLnInfo as l inner join tblSurveyRate AS r on l.Date_Key>=r.asOfDate
GROUP BY ProductCode, Int(4*(100*[InterestRate]-IIf(NumberOfMonths<240,r.FN15,r.FN30)))*25+25
ORDER BY ProductCode, Int(4*(100*[InterestRate]-IIf(NumberOfMonths<240,r.FN15,r.FN30)))*25+25;



'------ qryHistProdInfo  20140603
SELECT ProductCode, Sum(curBal)/1000000 AS chBal, Sum(Rate*curBal)/1000000/chBal AS GNR, Sum(incBps*curBal)/1000000/chBal AS incentive, 
	Sum(age*curBal)/1000000/chBal AS WALA, Sum(FICO*curBal)/1000000/chBal AS oFICO, Sum(LTV*curBal)/1000000/chBal AS oLTV, Sum(curLTV*curBal)/1000000/chBal AS cLTV
FROM qryMonthEndChLnInfo
GROUP BY ProductCode;


'------ qryMonthEndChLnInfo --- 20140603
SELECT l.LoanNumber, l.FundingDate, l.LoanAmount, l.PFIOwner, p.ParticipationPercent AS chicagoParticipation, PIAmount, 100*[InterestRate] AS rate, 
	iif(isnull(Sched_End_Prin_Bal),fundedAmount, Sched_End_Prin_Bal)*p.ParticipationPercent AS curBal, Year([FundingDate]) AS oYear, l.ProductCode, 
	100*l.LoanToValue as LTV, l.PropertyState, l.TotalLoanToValue, l.OrigLoanAmount, 100*[CurrentLoanToValue] AS curLTV, l.Sched_End_Prin_Bal,
	l.NumberOfMonths, DateDiff("m",[fundingDate],l.Date_Key) AS age, l.Date_Key, 100*(100*[InterestRate]-IIf(NumberOfMonths<240,r.FN15,r.FN30)) AS incBps,
	Int(4*(100*[InterestRate]-IIf(NumberOfMonths<240,r.FN15,r.FN30)))*25+25 AS incUpper, l.FirstPaymentDate,
	IIf([FICOScore] Is Null , 
		IIf ( IsNumeric([CoBorrower1FICOScore]), [CoBorrower1FICOScore] , 620.0 ),
	IIf( IsNumeric([FICOScore]) ,
		IIf ( [CoBorrower1FICOScore] Is Null , FICOScore, 
		IIf	( IsNumeric(CoBorrower1FICOScore), IIf ( CoBorrower1FICOScore>FICOScore , FICOScore , CoBorrower1FICOScore ), FICOScore )),
	IIf( IsNumeric(CoBorrower1FICOScore), CoBorrower1FICOScore, 620.0) )) AS FICO 
FROM  (dbo_UV_DM_MPF_Archive_Loan AS l INNER JOIN dbo_UV_DM_MPF_Daily_DeliveryCommitments_Participation AS p ON l.deliveryCommitmentNumber=p.deliveryCommitmentNumber)
inner join tblSurveyRate AS r on l.Date_Key>=r.asOfDate
WHERE l.Date_Key=[inDateKey]
and (iif(isnull(Sched_End_Prin_Bal),fundedAmount, Sched_End_Prin_Bal)>0 )
and   p.ParticipationOrgKey=3


'------ mkTmpLnInfo --- 20140603 for qryIncBal and qryCurLTV
SELECT l.LoanNumber, l.FundingDate, l.LoanAmount, p.ParticipationPercent AS chicagoParticipation, InterestRate , 
	iif(isnull(Sched_End_Prin_Bal),fundedAmount, Sched_End_Prin_Bal)*p.ParticipationPercent AS curBal, l.ProductCode, 
	l.PropertyState, l.OrigLoanAmount, l.Sched_End_Prin_Bal, 100*l.LoanToValue as LTV,
	l.NumberOfMonths, l.Date_Key, 
	l.FirstPaymentDate
into tmpLnInfo
FROM  (dbo_UV_DM_MPF_Archive_Loan AS l INNER JOIN dbo_UV_DM_MPF_Daily_DeliveryCommitments_Participation AS p ON l.deliveryCommitmentNumber=p.deliveryCommitmentNumber)
WHERE l.Date_Key=[inDateKey]
and (iif(isnull(Sched_End_Prin_Bal),fundedAmount, Sched_End_Prin_Bal)>0 )
and   p.ParticipationOrgKey=3


'========================== reference
'-- qryLoanCusip
SELECT "MPF" AS Portfolio, 
	IIf(Left([ProductCode],2)="MS","MS",IIf(Left([ProductCode],2)="GL","GL",IIf([RemittanceTypeID]=1,"SS",IIf([RemittanceTypeID]=3,"MA","AA")))) AS ScheduleType, 
	IIf([PortfolioIndicator]="BATCH",Year([OriginalLoanClosingDate]),Year([LoanRecordCreationDate])) AS OriginationYear, 
	l.LoanNumber, Right([ProductCode],2) AS AccountType, 
	(CInt([InterestRate]*200)/200)*100 AS PassThruRate, 
	[ScheduleType] & [OriginationYear] & [AccountType] & ([PassThruRate]*100) AS Cusip, 
	[OriginationYear] AS AccountClass, [MPFBalance]*[ChicagoParticipation] AS Notional, 
	IIf([CurrentLoanBalance]=0,[OriginalAmount]/[OriginalAmount],[CurrentLoanBalance]/[OriginalAmount]) AS Factor, 
	(CInt([InterestRate]*100000)/100000)*100 AS Wac, (CInt([Coupon]*100000)/100000)*100 AS Coup, 
	CInt(NPer(([InterestRate]/12),([PIAmount]*-1),[CurrentLoanBalance])) AS WAM, Age, 
	IIf([AccountType]=15,160,IIf([AccountType]=20,220,IIf([AccountType]=30,335,""))) AS Swam, EntryDate AS Settle, 
	IIf([ScheduleType]="GL","GNMA",IIf([ScheduleType]="MS","GNMA","FNMA")) AS Agency, 
	IIf([ScheduleType]="MS",18,IIf([ScheduleType]="GL",18,IIf([ScheduleType]="MA",2,IIf([ScheduleType]="AA",48,IIf([ScheduleType]="SS",18,18))))) AS Delay, 
	CurrentLoanBalance*ChicagoParticipation AS cBal, OriginalAmount, ChicagoParticipation, l.CEFee, l.DeliveryCommitmentNumber, PIAmount, interestRate
FROM tblMPFLoansAggregateSource AS l;

'---- qryLnInfo
SELECT LoanNumber, NumberOfMonths, ProductCode, FundingDate AS origDate, Year(FundingDate) AS origYear, FirstPaymentDate, 
   PIAmount, 100*[InterestRate] AS rate, 
   CDbl(IIf([FICOScore] Is Null, IIf(IsNumeric([CoBorrower1FICOScore]),[CoBorrower1FICOScore],620),
		IIf(IsNumeric([FICOScore]),IIf([CoBorrower1FICOScore] Is Null,FICOScore,
								IIf(IsNumeric(CoBorrower1FICOScore),IIf(CoBorrower1FICOScore>FICOScore,FICOScore,CoBorrower1FICOScore),FICOScore)),
		IIf(IsNumeric(CoBorrower1FICOScore),CoBorrower1FICOScore,620)))) AS FICO, 
   PropertyState, 100*[LoanToValue] AS LTV, LoanAmount, ParticipationPercent*[Sched_End_Prin_Bal] AS curBal, PFIOwner, 
   ParticipationPercent as ChicagoParticipation, 100*[CurrentLoanToValue] AS curLTV, Sched_End_Prin_Bal, 
   DateDiff("m",origDate,[r.asOfDate]) AS age, r.asOfDate, 100*(100*[InterestRate]-IIf(NumberOfMonths<240,r.FN15,r.FN30)) AS incBps, 
   Int(4*(100*[InterestRate]-IIf(NumberOfMonths<240,r.FN15,r.FN30)))*25+25 AS incUpper
FROM (dbo_UV_DM_MPF_Daily_Loan AS l
LEFT JOIN (SELECT DeliveryCommitmentNumber, ParticipationPercent
	FROM dbo_UV_DM_MPF_Daily_DeliveryCommitments_Participation 
	WHERE ParticipationOrgKey=3 ) AS p ON l.deliveryCommitmentNumber=p.deliveryCommitmentNumber) inner join qrySurveyRate AS r on l.LoanNumber>r.FN15
WHERE  l.Sched_End_Prin_Bal>0;


'-------- qryCurRate
SELECT *
FROM   tblSurveyRateHist
WHERE  asOfDate=(select max(asOfDate) from tblSurveyRateHist);

'-------- qrySurveyRate
SELECT *
FROM tblSurveyRate
WHERE (((tblSurveyRate.asOfDate)=[inDate]));


'======================== misc
'----- Open Basis from Bill Page
SELECT HLFAsOfDate, Count(HLFLoanNbr) AS CountOfLoanNumber, 
   Sum(HLFFHLBCDailyAssetPLAmt) AS SumOfFHLBCDailyAssetPL, 
   Sum(HLFLoanPaydownAdjAmt) AS SumOfPaydownAdjustment, 
   Sum(HLFFHLBCDailyAssetPLSumAmt) AS SumOfFHLBCDailyAssetPLSummary, 
   HLFHedgeID, Sum(HLFFHLBCIncepPLAmt) AS SumOfFHLBCIncepPL, Sum(HLFFHLBCDailyAssetPLEffectiveSumAmt) AS SumOfFHLBCDailyAssetPLEffectiveSummary
FROM [Historical HAUS MPF Basis Table]
GROUP BY HLFAsOfDate, HLFHedgeID, HLFEffectivenessInd
HAVING (((HLFAsOfDate)=[AsOfDate]) AND ((HLFEffectivenessInd)="1")) OR (((HLFAsOfDate)=[AsOfDate]) AND ((Sum(HLFLoanPaydownAdjAmt))<>0) AND ((HLFHedgeID) Is Null) AND ((HLFEffectivenessInd)<>"1"))
ORDER BY HLFAsOfDate;

'--------- qryOBA : loan level
SELECT HLFLoanNbr, HLFFHLBCIncepPLAmt, HLFLoanPaydownAdjAmt, HLFEffectivenessInd, HLFHedgeID
FROM dbo_HAUS_Loan_Fact
WHERE ((HLFAsOfDate)=[asOfDate]) And ((HLFEffectivenessInd)="1" Or (HLFLoanPaydownAdjAmt<>0 And (HLFHedgeID Is Null)));

'------- qryStripByDate -- get AF and ClosedBasis of cohort level
SELECT a.HAIAmortCode, Sum(a.HAIUnamortizedNetAmt) AS netAmt, q.OriginationYear, q.AccountType AS Term, q.PassThruRate AS GNR, q.Agency
FROM dbo_arch_Amortization_Information AS a INNER JOIN qryLoanCusip AS q ON a.HAILoanNbr=q.LoanNumber
WHERE (((a.HAIasOfDate)=[asOfDate]))
GROUP BY a.HAIAmortCode, q.OriginationYear, q.AccountType, q.PassThruRate, q.Agency
HAVING (((a.HAIAmortCode)<>"D"))
ORDER BY a.HAIAmortCode;

'----- qryLoanStrip
SELECT q.LoanNumber, a.HAIAmortCode, a.HAIUnamortizedNetAmt AS netAmt, q.OriginationYear, q.AccountType AS Term, q.Agency, q.Wac AS GNR, q.PassThruRate AS RateBucket, q.WAM, q.Age, q.cBal
FROM tblLoanStrip AS a INNER JOIN qryLoanCusip AS q ON a.HAILoanNbr = q.LoanNumber
ORDER BY a.HAIAmortCode;

'----- qryAllStrip
SELECT a.HAIAmortCode, Sum(a.HAIUnamortizedNetAmt) AS netAmt 
FROM tblLoanStrip AS a  
GROUP BY a.HAIAmortCode 
ORDER BY a.HAIAmortCode;

'------- qryLnIncBal
SELECT LoanNumber, ProductCode, [ChicagoParticipation]*[Sched_End_Prin_Bal] AS curBal, [InterestRate], 
   100*(100*[InterestRate]-IIf(ProductCode like "*15", [c15], [c30])) AS incBps, 
   Int(4*(100*[InterestRate]-IIf(ProductCode like "*15", [c15], [c30])) )*25 +25 AS incUpper
FROM dbo_UV_DM_MPF_Daily_Loan
WHERE ChicagoParticipation>0
and      Sched_End_Prin_Bal>0;

'------ qryFICOhist
SELECT h.Year_Key, q.Cusip, Count(q.LoanNumber) AS CountOfLoanNumber, Sum(q.cBal) AS curBal, Sum(q.cBal*iif(isNumeric(h.BORR_FICO), h.Borr_FICO,600))/curBal AS curFICO
FROM dbo_UV_FRA_MGIC_Final_Historical as h INNER JOIN qryLoanCusip as q ON h.LoanNumber = q.LoanNumber
WHERE (((h.Quarter_Key)=2))
GROUP BY h.Year_Key, q.Cusip
ORDER BY h.Year_Key;

'------ qryHistLnInfo
SELECT LoanNumber, NumberOfMonths, ProductCode, DateAdd("m",-1,[FirstPaymentDate]) AS origDate, Year(origDate) AS origYear, FirstPaymentDate, PIAmount, 100*[InterestRate] AS rate, 
	CDbl(IIf([FICOScore] Is Null,IIf([CoBorrower1FICOScore] Is Null,600,IIf(IsNumeric([CoBorrower1FICOScore]),[CoBorrower1FICOScore],600)),IIf(IsNumeric([FICOScore]),[FICOScore],600))) AS FICO, 
	PropertyState, 100*[LoanToValue] AS LTV, LoanAmount, [ChicagoParticipation]*[Sched_End_Prin_Bal] AS curBal, PFIOwner, ChicagoParticipation, 100*[CurrentLoanToValue] AS curLTV, 
	Sched_End_Prin_Bal, DateDiff("m",origDate,[r.asOfDate]) AS age, r.asOfDate, 100*(100*[InterestRate]-IIf(NumberOfMonths<240,r.FN15,r.FN30)) AS incBps, 
	Int(4*(100*[InterestRate]-IIf(NumberOfMonths<240,r.FN15,r.FN30)))*25+25 AS incUpper
FROM dbo_UV_DM_MPF_archive_Loan AS l, qrySurveyRate AS r
WHERE l.Date_key=[monthEndDt] 
and l.ChicagoParticipation>0 
And l.Sched_End_Prin_Bal>0;





			
	CAST(
			CASE 
			WHEN [FICOScore] Is Null THEN 
				CASE WHEN IsNumeric([CoBorrower1FICOScore])=1 THEN [CoBorrower1FICOScore] ELSE 620.0 END
			WHEN IsNumeric([FICOScore])=1 THEN
				CASE WHEN [CoBorrower1FICOScore] Is Null THEN FICOScore 
					WHEN IsNumeric(CoBorrower1FICOScore)=1 THEN 
						CASE WHEN CoBorrower1FICOScore>FICOScore THEN FICOScore ELSE CoBorrower1FICOScore END
					ELSE FICOScore 
				END
			WHEN IsNumeric(CoBorrower1FICOScore)=1 THEN CoBorrower1FICOScore
			ELSE 620.0
			END AS REAL) AS FICO 
			



'======================== backup
'---- qryLnInfo
'-- SELECT LoanNumber, NumberOfMonths, ProductCode, DateAdd("m",-1,[FirstPaymentDate]) AS origDate, Year(origDate) AS origYear, FirstPaymentDate, PIAmount, 100*[InterestRate] AS rate, 
'--    CDbl(IIf([FICOScore] Is Null,IIf([CoBorrower1FICOScore] Is Null,600,IIf(IsNumeric([CoBorrower1FICOScore]),[CoBorrower1FICOScore],600)),IIf(IsNumeric([FICOScore]),[FICOScore],600))) AS FICO, 
'--    PropertyState, 100*[LoanToValue] AS LTV, LoanAmount, [ChicagoParticipation]*[Sched_End_Prin_Bal] AS curBal, PFIOwner, ChicagoParticipation, 100*[CurrentLoanToValue] AS curLTV, 
'--    Sched_End_Prin_Bal, DateDiff("m", origDate, [asOfDate]) as age
'-- FROM dbo_UV_DM_MPF_Daily_Loan;

'---- qryLnInfo
'-- SELECT LoanNumber, NumberOfMonths, ProductCode, DateAdd("m",-1,[FirstPaymentDate]) AS origDate, Year(origDate) AS origYear, FirstPaymentDate, PIAmount, 100*[InterestRate] AS rate, 
'--    CDbl(IIf([FICOScore] Is Null,IIf([CoBorrower1FICOScore] Is Null,600,IIf(IsNumeric([CoBorrower1FICOScore]),[CoBorrower1FICOScore],600)),IIf(IsNumeric([FICOScore]),[FICOScore],600))) AS FICO, 
'--    PropertyState, 100*[LoanToValue] AS LTV, LoanAmount, [ChicagoParticipation]*[Sched_End_Prin_Bal] AS curBal, PFIOwner, ChicagoParticipation, 100*[CurrentLoanToValue] AS curLTV,   
'--    Sched_End_Prin_Bal, DateDiff("m", origDate,[r.asOfDate]) AS age, r.asOfDate,
'--    100*(100*[InterestRate]-IIf(ProductCode like "*15", r.FN15, r.FN30)) AS incBps, 
'--    Int(4*(100*[InterestRate]-IIf(ProductCode like "*15", r.FN15, r.FN30)) )*25 +25 AS incUpper
'-- FROM dbo_UV_DM_MPF_Daily_Loan as l, qryCurRate as r
'-- WHERE  l.ChicagoParticipation>0
'-- and    l.Sched_End_Prin_Bal>0

'--  qryLnCurLTV
'--   Select m.loanNumber, m.curBal, m.Sched_End_Prin_Bal*(m.LTV/m.LoanAmount)*(o.HPI/c.HPI) as curLTV, 
'--      Year(DateAdd("m", -4, FirstPaymentDate))&"M"&Month(DateAdd("m", -4, FirstPaymentDate))
'--   FROM   qryLnInfo m, dbo_FMHPIstate o, dbo_FMHPIstate c
'--   WHERE  c.yrMon="2010M12"   
'--   AND    o.state=c.state
'--   AND    m.PropertyState=o.state
'--   AND    o.yrMon=Year(DateAdd("m", -4, FirstPaymentDate))&"M"&IIf(Month(DateAdd("m", -4, FirstPaymentDate))<10,"0","")&Month(DateAdd("m", -4, FirstPaymentDate))

'-- SELECT m.loanNumber, m.curBal, m.Sched_End_Prin_Bal*(m.LTV/m.LoanAmount)*(o.HPI/c.HPI) AS curLTV, m.productCode
'-- FROM qryLnInfo AS m, dbo_FMHPIstate AS o, dbo_FMHPIstate AS c
'-- WHERE c.yrMon="2010M12" And o.state=c.state And m.PropertyState=o.state 
'-- And o.yrMon=Year(DateAdd("m",-4,FirstPaymentDate)) & "M" & IIf(Month(DateAdd("m",-4,FirstPaymentDate))<10,"0","") & Month(DateAdd("m",-4,FirstPaymentDate));

