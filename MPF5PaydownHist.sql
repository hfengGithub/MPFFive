if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt_MPF5PaydownHist]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[rpt_MPF5PaydownHist]
GO

CREATE TABLE [dbo].[rpt_MPF5PaydownHist] (
	[cusip] [varchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[period] [int] NULL ,
	[curBal] [float] NOT NULL ,
	[schPrincipal] [real] NOT NULL ,
	[asOfDate] [datetime] NOT NULL ,
	[SMM] [real] NULL ,
	[CPR1] [real] NULL ,
	[CPR3] [real] NULL ,
	[CPR6] [real] NULL ,
	[CPR12] [real] NULL 
) ON [PRIMARY]
GO


--------- the history of relationship between loan and cusip
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[rpt_uv_MPF5RefHist]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[rpt_uv_MPF5RefHist]
GO

CREATE VIEW dbo.uv_MPF5RefHist
AS
SELECT     cusip, loanNumber, asOfDate
FROM         research.dbo.MPF5RefHist
GO


TRUNCATE TABLE rpt_MPF5PaydownHist

INSERT INTO rpt_MPF5PaydownHist (cusip, curBal, schPrincipal, asOfDate)
SELECT r.cusip, SUM(h.ScheduleEndPrincipalBal*m.mdl_chicagoParticipation) AS curBal, 
   SUM(h.CurrentSchedulePrincipal*m.mdl_chicagoParticipation) AS schPrincipal, r.asOfDate
FROM   prepayRpt.dbo.rpt_uv_MPF5RefHist r INNER JOIN 
   prepayRpt.dbo.uv_ma_dc_loan m ON r.loanNumber=m.mdl_loanNumber INNER JOIN
   FHLB.dbo.FHLBHistory h ON 
      h.LoanNumber = r.LoanNumber AND 
      YEAR(h.PayDate) = YEAR(r.asOfDate) AND 
      MONTH(h.PayDate) = MONTH(r.asOfDate)
GROUP BY r.cusip, r.asOfDate

------ set period minDate==>0
DECLARE @d_minDate datetime
SELECT @d_minDate=MIN(asOfDate) FROM rpt_MPF5PaydownHist

update rpt_MPF5PaydownHist
set period=datediff(m, @d_minDate, asOfDate)

------ Set SMM
   UPDATE h2	
   SET    SMM=CASE WHEN 1-h2.curBal/(h1.curBal-h2.schPrincipal)>0 THEN 1-h2.curBal/(h1.curBal-h2.schPrincipal) ELSE 0 END
   FROM   rpt_MPF5PaydownHist h1, rpt_MPF5PaydownHist h2
   WHERE  h1.cusip=h2.cusip
   AND    h1.period+1=h2.period  

------ Set CPR1
   UPDATE rpt_MPF5PaydownHist	
   SET    CPR1 = 100*(1-Power((1-SMM), 12))

------ Set LTV = 100*Sum(origLTV*curBal)/sum(curBal)


------ Set CPR3
   UPDATE h3	
   SET    CPR3=100*(1-Power((1-h1.SMM)*(1-h2.SMM)*(1-h3.SMM), 4))
   FROM   rpt_MPF5PaydownHist h1, rpt_MPF5PaydownHist h2, rpt_MPF5PaydownHist h3
   WHERE  h1.cusip=h2.cusip
   AND    h2.cusip=h3.cusip
   AND    h1.period+1=h2.period  
   AND    h2.period+1=h3.period

------ Set CPR6
   UPDATE h6	
   SET    CPR6=100*(1-Power((1-h1.SMM)*(1-h2.SMM)*(1-h3.SMM)*(1-h4.SMM)*(1-h5.SMM)*(1-h6.SMM), 2))
   FROM   rpt_MPF5PaydownHist h1, rpt_MPF5PaydownHist h2, rpt_MPF5PaydownHist h3, rpt_MPF5PaydownHist h4, 
          rpt_MPF5PaydownHist h5, rpt_MPF5PaydownHist h6
   WHERE  h1.cusip=h2.cusip
   AND    h2.cusip=h3.cusip
   AND    h4.cusip=h3.cusip
   AND    h4.cusip=h5.cusip
   AND    h5.cusip=h6.cusip
   AND    h1.period+1=h2.period  
   AND    h2.period+1=h3.period
   AND    h3.period+1=h4.period  
   AND    h4.period+1=h5.period
   AND    h5.period+1=h6.period  

------ Set CPR12
   UPDATE h12	
   SET    CPR12=100*(1-(1-h1.SMM)*(1-h2.SMM)*(1-h3.SMM)*(1-h4.SMM)*(1-h5.SMM)*(1-h6.SMM)*
                       (1-h7.SMM)*(1-h8.SMM)*(1-h9.SMM)*(1-h10.SMM)*(1-h11.SMM)*(1-h12.SMM))
   FROM   rpt_MPF5PaydownHist h1, rpt_MPF5PaydownHist h2, rpt_MPF5PaydownHist h3, rpt_MPF5PaydownHist h4, 
          rpt_MPF5PaydownHist h5, rpt_MPF5PaydownHist h6, rpt_MPF5PaydownHist h7, rpt_MPF5PaydownHist h8,
          rpt_MPF5PaydownHist h9, rpt_MPF5PaydownHist h10, rpt_MPF5PaydownHist h11, rpt_MPF5PaydownHist h12
   WHERE  h1.cusip=h2.cusip
   AND    h2.cusip=h3.cusip
   AND    h3.cusip=h4.cusip
   AND    h4.cusip=h5.cusip
   AND    h5.cusip=h6.cusip
   AND    h6.cusip=h7.cusip
   AND    h7.cusip=h8.cusip
   AND    h8.cusip=h9.cusip
   AND    h9.cusip=h10.cusip
   AND    h10.cusip=h11.cusip
   AND    h11.cusip=h12.cusip
   AND    h1.period+1=h2.period  
   AND    h2.period+1=h3.period
   AND    h3.period+1=h4.period  
   AND    h4.period+1=h5.period
   AND    h5.period+1=h6.period  
   AND    h6.period+1=h7.period  
   AND    h7.period+1=h8.period
   AND    h8.period+1=h9.period  
   AND    h9.period+1=h10.period
   AND    h10.period+1=h11.period  
   AND    h11.period+1=h12.period  


---------------------------------- 20080320 version in Access
-- SELECT DateDiff("m",#1/1/2000#,asOfDate) AS Expr1
-- FROM LoanHist;

-- update tabA as A, tabB as B
-- set B.name=A.name
-- where A.id=B.id

INSERT INTO t2 (uname, rank)
SELECT uname, rank
FROM (Select * from t1
where  rank=1
union
Select * from t1
where  rank=2
)  AS A


---------- appLoanHist -------------- 1st step
INSERT INTO LoanHist (cusip, LoanNumber, asOfDate, curBal, schPrincipal )
SELECT cusip, LoanNumber, settle, cBal, 
   IIF(ChicagoParticipation*PIamount-cBal*interestRate/12>0, 
      IIF(ChicagoParticipation*PIamount-cBal*interestRate/12>cBal, cBal, ChicagoParticipation*PIamount-cBal*interestRate/12), 0)
FROM qryLoanCusip;

-- INSERT INTO paydownHist ( cusip, asOfDate, period, curBal, schPrincipal )
-- SELECT cusip, asOfDate, DateDiff("m",#1/1/2000#,asOfDate), sum(curBal), sum(schPrincipal)
-- FROM loanHist
-- GROUP BY cusip, asOfDate, DateDiff("m",#1/1/2000#,asOfDate);

------- appPaydownHist 
INSERT INTO paydownHist(cusip, asOfDate, period, curBal, nextBal, schPrincipal)
SELECT h1.cusip, h1.asOfDate, DateDiff("m", #1/1/2000#, h1.asOfDate) , sum(h1.curBal), SUM(h2.curBal), sum(h1.schPrincipal)
FROM  ( loanHist h1 LEFT JOIN loanHist h2  
   ON     h1.loanNumber=h2.loanNumber
   AND    DateDiff("m", h1.asOfDate, h2.asOfDate)=1 )
INNER JOIN qryMaxDate ON  h1.asOfDate < qryMaxDate.MaxHistDate
GROUP BY h1.cusip, h1.asOfDate, DateDiff("m", #1/1/2000#, h1.asOfDate)

------- appPaydownHistByDate ------ 2nd step
-- INSERT INTO paydownHist(cusip, asOfDate, period, curBal, nextBal, schPrincipal)
-- SELECT h1.cusip, h1.asOfDate, DateDiff("m", #1/1/2000#, h1.asOfDate) , sum(h1.curBal), SUM(h2.curBal), sum(h1.schPrincipal)
-- FROM  ( loanHist h1 LEFT JOIN loanHist h2  
--    ON     h1.loanNumber=h2.loanNumber
--    AND    DateDiff("m", h1.asOfDate, h2.asOfDate)=1 )
-- INNER JOIN qryMaxDate ON  h1.asOfDate < qryMaxDate.MaxHistDate
-- WHERE  h1.asOfDate=[periodBeginDate]
-- GROUP BY h1.cusip, h1.asOfDate, DateDiff("m", #1/1/2000#, h1.asOfDate)

------- appPaydownHistByDate ------ 2nd step ----- the default values of SMM, CPR1, 3, 6, 12 are all 0
-- Hua 20150623 -- Change to load 1 month data, no need for the qryMaxDate. Make sure the input periodBeginDate is correct.
INSERT INTO paydownHist(cusip, asOfDate, period, curBal, nextBal, schPrincipal)
SELECT h1.cusip, h1.asOfDate, DateDiff("m", #1/1/2000#, h1.asOfDate) , sum(h1.curBal), SUM(h2.curBal), sum(h1.schPrincipal)
FROM   loanHist h1 LEFT JOIN loanHist h2 ON h1.loanNumber=h2.loanNumber AND DateDiff("m", h1.asOfDate, h2.asOfDate)=1 
WHERE  h1.asOfDate=[periodBeginDate]
GROUP BY h1.cusip, h1.asOfDate, DateDiff("m", #1/1/2000#, h1.asOfDate)

----- qryMaxDate
SELECT Max(LoanHist.asOfDate) AS MaxHistDate, Min(LoanHist.asOfDate) AS MinOfasOfDate
FROM LoanHist;


------ appMPF_ALL --------------    3rd step
-- Hua 20160120 -- not counting the GNMBS loans -- ie. cusip like GL????0*
-- INSERT INTO paydownHist ( cusip, asOfDate, period, curBal, nextBal, schPrincipal )
-- SELECT "MPF_ALL", [periodBeginDate], period, sum(curBal), sum(nextBal), sum(schPrincipal)
-- FROM paydownHist
-- WHERE cusip Not Like "MPF_*" 
-- AND cusip not like 'GL????0*' 
-- AND asOfDate=[periodBeginDate]
-- GROUP BY period;


------ appMPF_ALL --------------    3rd step
-- Hua 20160120 -- not counting the GNMBS loans -- ie. cusip like GL????0*
-- Hua 20160322 -- skipping the GNMBS jumbo loans  
INSERT INTO paydownHist ( cusip, asOfDate, period, curBal, nextBal, schPrincipal )
SELECT "MPF_ALL", [periodBeginDate], period, sum(curBal), sum(nextBal), sum(schPrincipal)
FROM paydownHist
WHERE cusip Not Like "MPF_*" 
AND mid(cusip,7,2) not IN ("03","43","05","45") 
AND asOfDate=[periodBeginDate]
GROUP BY period;


---------- use DSum() to reset the summary row ====================================================================================
-- update paydownHist 
-- set schPrincipal=DSum(schPrincipal, "paydownHist", "cusip Not Like 'MPF_*' and cusip not like 'GL????0*' And asOfDate=#12/18/2015#"  )
-- where asOfDate=#12/18/2015#
-- AND   cusip="MPF_ALL"


------ Set SMM
   UPDATE paydownHist	
   SET    SMM=IIF(1-nextBal/(curBal-schPrincipal)>0 , 1-nextBal/(curBal-schPrincipal), 0 )
   WHERE  asOfDate=[periodBeginDate]

------ Set CPR1
   UPDATE paydownHist SET CPR1 = 100*(1-(1-SMM)^12 );
   WHERE asOfDate=[periodBeginDate]
   -- UPDATE paydownHist	
   -- SET    CPR1 = 100*(1-exp(12*log(1-SMM)))

------ Set CPR3
   -- SET    h3.CPR3 = 100*(1-exp(4*log((1-h1.SMM)*(1-h2.SMM)*(1-h3.SMM))))

   UPDATE paydownHist AS h1, paydownHist AS h2, paydownHist AS h3, 
   SET    h3.CPR3 = 100*(1-((1-h1.SMM)*(1-h2.SMM)*(1-h3.SMM))^4)
   WHERE  h3.asOfDate=[periodBeginDate]
   AND    h1.cusip=h2.cusip
   AND    h2.cusip=h3.cusip
   AND    h1.period+1=h2.period  
   AND    h2.period+1=h3.period
-- 
-- SELECT asOfDate, 1- sum(nextBal)/(sum(curBal)-sum(schPrincipal)) AS SMM,
--    100*(1-exp(12*log(1-SMM))) AS CPR1
-- FROM PaydownHist
-- GROUP BY asOfDate
-- 
------ Set CPR6
   UPDATE paydownHist AS h1, paydownHist AS h2, paydownHist AS h3, paydownHist AS h4, paydownHist AS h5, paydownHist AS h6
   SET    h6.CPR6 = 100*(1-((1-h1.SMM)*(1-h2.SMM)*(1-h3.SMM)*(1-h4.SMM)*(1-h5.SMM)*(1-h6.SMM))^2)
   WHERE  h6.asOfDate=[periodBeginDate]
   AND    h1.cusip=h2.cusip
   AND    h2.cusip=h3.cusip
   AND    h3.cusip=h4.cusip
   AND    h4.cusip=h5.cusip
   AND    h5.cusip=h6.cusip
   AND    h1.period+1=h2.period  
   AND    h2.period+1=h3.period
   AND    h3.period+1=h4.period  
   AND    h4.period+1=h5.period
   AND    h5.period+1=h6.period  

------ Set CPR12   
   UPDATE paydownHist AS h1, paydownHist AS h2, paydownHist AS h3, paydownHist AS h4, paydownHist AS h5, paydownHist AS h6,
	paydownHist AS h7, paydownHist AS h8, paydownHist AS h9, paydownHist AS h10, paydownHist AS h11, paydownHist AS h12
   SET    h12.CPR12=100*(1-(1-h1.SMM)*(1-h2.SMM)*(1-h3.SMM)*(1-h4.SMM)*(1-h5.SMM)*(1-h6.SMM)*
                       (1-h7.SMM)*(1-h8.SMM)*(1-h9.SMM)*(1-h10.SMM)*(1-h11.SMM)*(1-h12.SMM))
   WHERE  h12.asOfDate=[periodBeginDate]
   AND    h1.cusip=h2.cusip
   AND    h2.cusip=h3.cusip
   AND    h3.cusip=h4.cusip
   AND    h4.cusip=h5.cusip
   AND    h5.cusip=h6.cusip
   AND    h6.cusip=h7.cusip
   AND    h7.cusip=h8.cusip
   AND    h8.cusip=h9.cusip
   AND    h9.cusip=h10.cusip
   AND    h10.cusip=h11.cusip
   AND    h11.cusip=h12.cusip
   AND    h1.period+1=h2.period  
   AND    h2.period+1=h3.period
   AND    h3.period+1=h4.period  
   AND    h4.period+1=h5.period
   AND    h5.period+1=h6.period  
   AND    h6.period+1=h7.period  
   AND    h7.period+1=h8.period
   AND    h8.period+1=h9.period  
   AND    h9.period+1=h10.period
   AND    h10.period+1=h11.period  
   AND    h11.period+1=h12.period  
   
   
   
   
   
   
-------------- 20080408
-- Too slow
-- SELECT LoanHist.asOfDate, LoanHist.loanNumber
-- FROM LoanHist
-- WHERE ( (dateDiff("m", startDate, asOfDate)=0) and loanNumber not in (select loanNumber from loanHist where dateDiff("m", startDate, asOfDate)=1))

INSERT INTO tmpLoans (asOfDate, loanNumber)
SELECT asOfDate, loanNumber
FROM   LoanHist
WHERE  (dateDiff("m", startDate, asOfDate)=0) 

and loanNumber not in (select loanNumber from loanHist where dateDiff("m", startDate, asOfDate)=1))


select distinct loanNUmber, coupon into loan1  from 
[SELECT LoanNumber, Coupon
FROM tblMPFLoansAggregateSource20080123
union
select loanNUmber, coupon
from tblMPFLoansAggregateSource20080215
union
select loanNUmber, coupon
from tblMPFLoansAggregateSource20080318
]. as a



select distinct loanNUmber, coupon into loanMaster  from 
[SELECT LoanNumber, Coupon
FROM loan1
union
select loanNUmber, coupon
from tblMPFLoansAggregateSource20080318
]. as a


---- delete asOFDate=2/15/2008
delete FROM tmpLoans where loanNumber in (
SELECT loanNumber
FROM LoanHist
where asOfDate=[aDate]
)

----- All paidOff field
update  LoanHist as h1, LoanHist as h2
set h1.paidOff="N"
where  h1.loanNumber=h2.loanNumber
and     dateDiff("M",  h1.asOfDate, h2.asOfDate)=1


-- update  loanMaster INNER JOIN tblMPFLoansAggregateSource20080123 ON loanMaster.loanNUmber = tblMPFLoansAggregateSource20080123.LoanNumber
-- set loanMaster.GNR=tblMPFLoansAggregateSource20080123.interestRate*100


---------- insTmpMst
SELECT loanNUmber, coupon, PIAmount, DeliveryCommitmentNumber, MANumber, PFINumber, OriginalAmount, ProductCode, ChicagoParticipation, CEFee, CEPerformanceFee, ExcessServicingFee, ServicingFee
into   tmpMst
FROM tblMPFLoansAggregateSource20080123;

--------- appLoanMaster
INSERT INTO loanMaster (
   loanNUmber, coupon, interestRate, PIAmount, DeliveryCommitmentNumber, MANumber, PFINumber, OriginalAmount, ProductCode, 
   ChicagoParticipation, CEFee, CEPerformanceFee, ExcessServicingFee, ServicingFee
)
SELECT loanNUmber, coupon, interestRate, PIAmount, DeliveryCommitmentNumber, MANumber, PFINumber, OriginalAmount, ProductCode, 
   ChicagoParticipation, CEFee, CEPerformanceFee, ExcessServicingFee, ServicingFee
FROM tmpMst;

SELECT LoanHist.asOfDate AS startDate, Sum(LoanHist.curBal) AS SumOfcurBal, 100*Sum([curBal]*[LoanMaster.coupon])/Sum([curBal]) AS WACoupon, 100*Sum([curBal]*[interestRate])/Sum([curBal]) AS GNR
FROM (LoanHist INNER JOIN loanMaster ON LoanHist.loanNumber = loanMaster.loanNUmber) INNER JOIN tblMPFLoansAggregateSource20080123 ON loanMaster.loanNUmber = tblMPFLoansAggregateSource20080123.LoanNumber
WHERE (((LoanHist.paidOff) Is Null))
GROUP BY LoanHist.asOfDate
HAVING (((LoanHist.asOfDate)=[startDate]));


----------- updPaidOff
UPDATE LoanHist AS h1, LoanHist AS h2 SET h1.paidOff = "N"
WHERE h1.asOfDate=[lastDateBeforePayOff]
and    h1.loanNumber=h2.loanNumber
and    dateDiff("m",  lastDateBeforePayOff, h2.asOfDate)=1;


----------- loan size & LTV
SELECT LoanHist.cusip, Sum(LoanHist.curBal) AS cBal, Sum([loanMaster].[OriginalAmount]*[curBal])/[cBal]/1000 AS lnsz, Sum(loanMaster.origLTV*curBal)/cBal AS oLTV
FROM loanMaster INNER JOIN LoanHist ON loanMaster.loanNUmber = LoanHist.loanNumber
WHERE (((LoanHist.asOfDate)=[histDate]))
GROUP BY LoanHist.cusip;

----------------- 20080505 crosstab query
-- TRANSFORM Sum(IDCPrice.price*[Holding])/Sum([holding]) AS AWprice
-- SELECT IDCPrice.AsOfDate
-- FROM MBSCMO INNER JOIN IDCPrice ON MBSCMO.CUSIP = IDCPrice.CUSIP
-- WHERE (((MBSCMO.SubActII) Like "PRIVATE*"))
-- GROUP BY IDCPrice.AsOfDate
-- ORDER BY IDCPrice.AsOfDate
-- PIVOT MBSCMO.Account;


----- 20080603
SELECT dbo_AFTMaster.State, Count(qryLoanCusip20080418o.LoanNumber) AS Cnt, Sum(qryLoanCusip20080418o.cBal) AS CurBal, Sum(qryLoanCusip20080418o.Coup*cBal)/CurBal AS Coupon, 
   Sum(qryLoanCusip20080418o.Wac*cBal)/CurBal AS curWac, Sum(qryLoanCusip20080418o.WAM*cBal)/CurBal AS curWAM, Sum(qryLoanCusip20080418o.Age*cBal)/CurBal AS WALA, 
   Sum(qryLoanCusip20080418o.OriginalAmount*cBal)/CurBal AS oLnsz, LoanHist.paidOff
FROM (qryLoanCusip20080418o INNER JOIN LoanHist ON qryLoanCusip20080418o.LoanNumber = LoanHist.loanNumber) INNER JOIN 
   dbo_AFTMaster ON LoanHist.loanNumber = cdbl(dbo_AFTMaster.LoanIDNumber)
WHERE (((qryLoanCusip20080418o.Cusip) In ("MA200630700","MA200730700","MA200630650", "MA200730650")) and (LoanHist.asOfDate)=#4/18/2008# )
GROUP BY dbo_AFTMaster.State, LoanHist.paidOff 
ORDER BY dbo_AFTMaster.State, LoanHist.paidOff;


SELECT dbo_AFTMaster.State,cusip,  Count(qryLoanCusip20080418o.LoanNumber) AS Cnt, Sum(qryLoanCusip20080418o.cBal) AS CurBal, Sum(qryLoanCusip20080418o.Coup*cBal)/CurBal AS Coupon, Sum(qryLoanCusip20080418o.Wac*cBal)/CurBal AS curWac, Sum(qryLoanCusip20080418o.WAM*cBal)/CurBal AS curWAM, Sum(qryLoanCusip20080418o.Age*cBal)/CurBal AS WALA, Sum(qryLoanCusip20080418o.OriginalAmount*cBal)/CurBal AS oLnsz
FROM dbo_AFTMaster, qryLoanCusip20080418o INNER JOIN LoanHist ON qryLoanCusip20080418o.LoanNumber = LoanHist.loanNumber
WHERE (((qryLoanCusip20080418o.Cusip) In ("MA200630700","MA200730600","MA200630650","MA200730650")) and LoanHist.loanNumber = cdbl(dbo_AFTMaster.LoanIDNumber) and (LoanHist.asOfDate)=#4/18/2008# )
GROUP BY dbo_AFTMaster.State,cusip
ORDER BY dbo_AFTMaster.State,cusip;

---------- 20140715 qryPFICPR
SELECT h.Date_Key AS beginDate, l.PFIName, sum(h.curBal)/1000000 AS beginBal, 
	iif(sum(h.curBal-h.schPrincipal)>sum(h.nextBal),1-(sum(h.nextBal)/sum(h.curBal- h.schPrincipal))^12, 0)*100 AS CPR
FROM dbo_chMPFhist AS h INNER JOIN dbo_uv_MPFLoan AS l ON h.LoanNumber = l.LoanNumber
WHERE date_key>#5/1/2013#
GROUP BY h.date_key, l.PFIName
HAVING sum(h.curBal)/1000000>100
ORDER BY h.date_key, l.PFIName;

---- qryCusipPaydown 20160204 filter out GNMBS loans
SELECT asOfDate, cusip, SMM, CPR1, CPR3, CPR6, CPR12, curBal-nextBal AS paydown, curBal, nextBal, schPrincipal
FROM PaydownHist
WHERE (((asOfDate)=[startDate]) AND ((cusip)<>"MPF_ALL") and mid(cusip,7,2) not in ("03", "05"))
ORDER BY asOfDate, curBal DESC , cusip;


------- Rollback the wrong data loaded
DELETE *
FROM paydownHist
WHERE asOfDate=#3/18/2015#

DELETE *
FROM loanHist
where asOfDate=#4/17/2015#
