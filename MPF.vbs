
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

'--------- qryOBA
SELECT HLFLoanNbr, HLFFHLBCIncepPLAmt, HLFLoanPaydownAdjAmt, HLFEffectivenessInd, HLFHedgeID
FROM dbo_HAUS_Loan_Fact
WHERE ((HLFAsOfDate)=[asOfDate]) And ((HLFEffectivenessInd)="1" Or (HLFLoanPaydownAdjAmt<>0 And (HLFHedgeID Is Null)));


'------- qryStrip -- get AF and ClosedBasis
SELECT a.HAIAmortCode, Sum(a.HAIUnamortizedNetAmt) AS netAmt, q.OriginationYear, q.AccountType AS Term, q.PassThruRate AS GNR, q.Agency
FROM dbo_haus_amortization_information AS a INNER JOIN qryLoanCusip AS q ON a.HAILoanNbr = q.LoanNumber
WHERE (((a.HAIasOfDate)=[asOfDate]))
GROUP BY a.HAIAmortCode, q.OriginationYear, q.AccountType, q.PassThruRate, q.Agency
HAVING (((a.HAIAmortCode)<>"D"))
ORDER BY a.HAIAmortCode;

'------- qryStripByDate -- get AF and ClosedBasis
SELECT a.HAIAmortCode, Sum(a.HAIUnamortizedNetAmt) AS netAmt, q.OriginationYear, q.AccountType AS Term, q.PassThruRate AS GNR, q.Agency
FROM dbo_arch_Amortization_Information AS a INNER JOIN qryLoanCusip AS q ON a.HAILoanNbr=q.LoanNumber
WHERE (((a.HAIasOfDate)=[asOfDate]))
GROUP BY a.HAIAmortCode, q.OriginationYear, q.AccountType, q.PassThruRate, q.Agency
HAVING (((a.HAIAmortCode)<>"D"))
ORDER BY a.HAIAmortCode;


SELECT tblMPFLoansAggregateSource.LoanNumber, tblMPFLoansAggregateSource.LastFundingEntryDate, tblMPFLoansAggregateSource.InterestRate, tblMPFLoansAggregateSource.Coupon, tblMPFLoansAggregateSource.ProductCode, tblMPFLoansAggregateSource.RemittanceTypeID, tblMPFLoansAggregateSource.ChicagoParticipation*tblMPFLoansAggregateSource.CurrentLoanBalance AS cBal, tblMPFLoansAggregateSource.EntryDate, uv_haus_amortization_info.HAIAmortCode, uv_haus_amortization_info.HAIHedgeID, uv_haus_amortization_info.HAILoanBalanceAmt, uv_haus_amortization_info.HAIUnamortizedNetAmt
FROM tblMPFLoansAggregateSource INNER JOIN uv_haus_amortization_info ON tblMPFLoansAggregateSource.LoanNumber=uv_haus_amortization_info.HAILoanNbr
WHERE (((uv_haus_amortization_info.HAIAmortCode) In ("D","H","X","A")));



'---- qryLnInfo
'-- SELECT LoanNumber, NumberOfMonths, ProductCode, DateAdd("m",-1,[FirstPaymentDate]) AS origDate, Year(origDate) AS origYear, FirstPaymentDate, PIAmount, 100*[InterestRate] AS rate, 
'--    CDbl(IIf([FICOScore] Is Null,IIf([CoBorrower1FICOScore] Is Null,550,IIf(IsNumeric([CoBorrower1FICOScore]),[CoBorrower1FICOScore],550)),IIf(IsNumeric([FICOScore]),[FICOScore],550))) AS FICO, 
'--    PropertyState, 100*[LoanToValue] AS LTV, LoanAmount, [ChicagoParticipation]*[Sched_End_Prin_Bal] AS curBal, PFIOwner, ChicagoParticipation, 100*[CurrentLoanToValue] AS curLTV, 
'--    Sched_End_Prin_Bal, DateDiff("m", origDate, [asOfDate]) as age
'-- FROM dbo_UV_DM_MPF_Daily_Loan;

'-------- qryCurRate
SELECT *
FROM tblSurveyRate
WHERE asOfDate=(select max(asOfDate) from tblSurveyRate);

'---- qryLnInfo
'-- SELECT LoanNumber, NumberOfMonths, ProductCode, DateAdd("m",-1,[FirstPaymentDate]) AS origDate, Year(origDate) AS origYear, FirstPaymentDate, PIAmount, 100*[InterestRate] AS rate, 
'--    CDbl(IIf([FICOScore] Is Null,IIf([CoBorrower1FICOScore] Is Null,550,IIf(IsNumeric([CoBorrower1FICOScore]),[CoBorrower1FICOScore],550)),IIf(IsNumeric([FICOScore]),[FICOScore],550))) AS FICO, 
'--    PropertyState, 100*[LoanToValue] AS LTV, LoanAmount, [ChicagoParticipation]*[Sched_End_Prin_Bal] AS curBal, PFIOwner, ChicagoParticipation, 100*[CurrentLoanToValue] AS curLTV,   
'--    Sched_End_Prin_Bal, DateDiff("m", origDate,[r.asOfDate]) AS age, r.asOfDate,
'--    100*(100*[InterestRate]-IIf(ProductCode like "*15", r.FN15, r.FN30)) AS incBps, 
'--    Int(4*(100*[InterestRate]-IIf(ProductCode like "*15", r.FN15, r.FN30)) )*25 +25 AS incUpper
'-- FROM dbo_UV_DM_MPF_Daily_Loan as l, qryCurRate as r
'-- WHERE  l.ChicagoParticipation>0
'-- and    l.Sched_End_Prin_Bal>0

'---- qryLnInfo
SELECT LoanNumber, NumberOfMonths, ProductCode, DateAdd("m",-1,[FirstPaymentDate]) AS origDate, Year(origDate) AS origYear, FirstPaymentDate, PIAmount, 
   100*[InterestRate] AS rate, 
   CDbl(IIf([FICOScore] Is Null,IIf([CoBorrower1FICOScore] Is Null,550,IIf(IsNumeric([CoBorrower1FICOScore]),[CoBorrower1FICOScore],550)),IIf(IsNumeric([FICOScore]),[FICOScore],550))) AS FICO, 
   PropertyState, 100*[LoanToValue] AS LTV, LoanAmount, [ChicagoParticipation]*[Sched_End_Prin_Bal] AS curBal, PFIOwner, ChicagoParticipation, 100*[CurrentLoanToValue] AS curLTV, 
   Sched_End_Prin_Bal, DateDiff("m",origDate,[r.asOfDate]) AS age, r.asOfDate, 
   100*(100*[InterestRate]-IIf(NumberOfMonths<240,r.FN15,r.FN30)) AS incBps, 
   Int(4*(100*[InterestRate]-IIf(NumberOfMonths<240,r.FN15,r.FN30)))*25+25 AS incUpper
FROM dbo_UV_DM_MPF_Daily_Loan AS l, qryCurRate AS r
WHERE l.ChicagoParticipation>0 And l.Sched_End_Prin_Bal>0;


'----- mkTblLoanStrip
SELECT HAILoanNbr, HAIAmortCode, HAIUnamortizedNetAmt 
into   tblLoanStrip
FROM dbo_arch_Amortization_Information  
WHERE (HAIAmortCode)<>"D"
AND   (HAIasOfDate)=[asOfDate]


'----- qryLoanStrip
SELECT q.LoanNumber, a.HAIAmortCode, a.HAIUnamortizedNetAmt AS netAmt, q.OriginationYear, q.AccountType AS Term, q.Agency, q.Wac AS GNR, q.PassThruRate AS RateBucket, q.WAM, q.Age, q.cBal
FROM tblLoanStrip AS a INNER JOIN qryLoanCusip AS q ON a.HAILoanNbr = q.LoanNumber
ORDER BY a.HAIAmortCode;

'----- qryAllStrip
SELECT a.HAIAmortCode, Sum(a.HAIUnamortizedNetAmt) AS netAmt 
FROM tblLoanStrip AS a  
GROUP BY a.HAIAmortCode 
ORDER BY a.HAIAmortCode;

'---- qryProdInfo
SELECT ProductCode, Sum(curBal)/1000000 AS chBal, Sum(Rate*curBal)/1000000/chBal AS GNR, Sum(incBps*curBal)/1000000/chBal AS incentive, Sum(age*curBal)/1000000/chBal AS WALA, 
   Sum(FICO*curBal)/1000000/chBal AS oFICO, Sum(LTV*curBal)/1000000/chBal AS oLTV, Sum(curLTV*curBal)/1000000/chBal AS cLTV
FROM qryLnInfo
GROUP BY ProductCode;


'--------- qryProdOBA
SELECT l.productCode, Sum(HLFFHLBCIncepPLAmt+HLFLoanPaydownAdjAmt)/1000000 AS OBA
FROM   tblOBA AS f INNER JOIN dbo_UV_DM_MPF_Daily_Loan AS l ON f.HLFLoanNbr = l.LoanNumber
WHERE  l.ChicagoParticipation>0
and    l.Sched_End_Prin_Bal>0
GROUP BY l.productCode;


'------- qryLnIncBal
SELECT LoanNumber, ProductCode, [ChicagoParticipation]*[Sched_End_Prin_Bal] AS curBal, [InterestRate], 
   100*(100*[InterestRate]-IIf(ProductCode like "*15", [c15], [c30])) AS incBps, 
   Int(4*(100*[InterestRate]-IIf(ProductCode like "*15", [c15], [c30])) )*25 +25 AS incUpper
FROM dbo_UV_DM_MPF_Daily_Loan
WHERE ChicagoParticipation>0
and      Sched_End_Prin_Bal>0;



=J3>(DATE(YEAR(TODAY())+1, MONTH(TODAY())+6, DAY(TODAY()))-TODAY()-3)


'--- get FX15 CPRs
SELECT  sum(curBal)/1000000 as beginBal, sum(cpr1*curBal)/sum(curBal) as waCPR1, sum(cpr3*curBal)/sum(curBal) as waCPR3, sum(cpr6*curBal)/sum(curBal) as waCPR6 
FROM PaydownHist 
WHERE MID(cusip,7,2)="15"
and asOfDate=#10/18/2011#
and left(cusip,2)<>"GL"


'-------- qryCurLTV 3 month delay
SELECT m.productCode, Sum(m.curBal*m.Sched_End_Prin_Bal*(m.LTV/m.LoanAmount)*(o.HPI/c.HPI))/Sum(m.curBal) AS curLTV
FROM qryLnInfo AS m, dbo_FMHPIstate AS o, dbo_FMHPIstate AS c
WHERE c.yrMon=(Select Max(yrMon) from dbo_FMHPIstate)   
   AND    o.state=c.state
   AND    m.PropertyState=o.state
   AND    o.yrMon=Year(DateAdd("m", -4, FirstPaymentDate))&"M"&IIf(Month(DateAdd("m", -4, FirstPaymentDate))<10,"0","")&Month(DateAdd("m", -4, FirstPaymentDate))
GROUP BY m.productCode;

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

