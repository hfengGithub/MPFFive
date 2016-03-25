
------- 20140102 Transact-SQL

--- uv_FNpxFICOLTV   +HighLTV

create view uv_FNpxFICOLTV as 
SELECT  loanNumber, curBal, origYear, PFInumber, FICOScore, LoanToValue,
	case when ficoscore>=740 then
		case 
		when LoanToValue<=.6 then -.0025
		when LoanToValue<=.75 then 0
		when LoanToValue<=.95 then .0025
		when LoanToValue<=.97 then .0075
		else 10000 end
	when ficoScore>=720 then 
		case 
		when LoanToValue<=.6 then -.0025
		when LoanToValue<=.70 then 0
		when LoanToValue<=.75 then 0.0025
		when LoanToValue<=.95 then .005
		when  LoanToValue<=.97 then .01
		else 10000 end
	when ficoScore>=700 then 
		case 
		when LoanToValue<=.6 then -.0025
		when LoanToValue<=.70 then 0.005
		when LoanToValue<=.75 then 0.0075
		when LoanToValue<=.97 then 0.01
		else 10000 end
	when ficoScore>=680 then 
		case 
		when LoanToValue<=.6 then 0
		when LoanToValue<=.70 then 0.005
		when LoanToValue<=.75 then 0.0125
		when LoanToValue<=.8 then 0.0175
		when LoanToValue<=.85 then 0.015
		when LoanToValue<=.95 then 0.0125
		when LoanToValue<=.97 then .015
		else 10000 end
	when ficoScore>=660 then 
		case 
		when LoanToValue<=.6 then 0
		when LoanToValue<=.70 then 0.01
		when LoanToValue<=.75 then 0.02
		when LoanToValue<=.8 then 0.025
		when LoanToValue<=.85 then 0.0275
		when LoanToValue<=.95 then 0.0225
		when LoanToValue<=.97 then .0225 
		else 10000 end
	when ficoScore>=640 then 
		case 
		when LoanToValue<=.6 then 0.005
		when LoanToValue<=.70 then 0.0125
		when LoanToValue<=.75 then 0.025
		when LoanToValue<=.8 then 0.03
		when LoanToValue<=.85 then 0.0325
		when LoanToValue<=.95 then 0.0275
		when LoanToValue<=.97 then .0275 
		else 10000 end
	when ficoScore>=620 then 
		case 
		when LoanToValue<=.6 then 0.005
		when LoanToValue<=.70 then 0.015
		when LoanToValue<=.8 then 0.03
		when LoanToValue<=.95 then 0.0325
		when LoanToValue<=.97 then .035 
		else 10000 end
	else 
		case 
		when LoanToValue<=.6 then 0.005
		when LoanToValue<=.70 then 0.015
		when LoanToValue<=.8 then 0.03
		when LoanToValue<=.95 then 0.0325
		when LoanToValue<=.97 then 0.0375 
		else 10000 end
	end AS px	
FROM MPFLoan



---- uv_FNpxCashout
create view uv_FNpxCashout as 
SELECT loanNumber, LoanPurpose, ficoScore, loanToValue,
	case when LoanPurpose='2' then 
		case
		when ficoscore>=740 then 
			case
			when LoanToValue<=.6 then 0
			when LoanToValue<=.75 then 0.0025
			when LoanToValue<=.80 then 0.005
			when LoanToValue<=.85 then 0.00625
			else 10000 end
		when ficoScore>=720 then
			case
			when LoanToValue<=.6 then 0
			when LoanToValue<=.75 then 0.00625
			when LoanToValue<=.80 then 0.0075
			when LoanToValue<=.85 then 0.015 
			else 10000 end
		when ficoScore>=700 then 
			case
			when LoanToValue<=.6 then 0
			when LoanToValue<=.75 then 0.00625
			when LoanToValue<=.80 then 0.0075
			when LoanToValue<=.85 then 0.015 
			else 10000 end
		when ficoScore>=680 then 
			case
			when LoanToValue<=.6 then 0
			when LoanToValue<=.75 then 0.0075
			when LoanToValue<=.80 then 0.01375
			when LoanToValue<=.85 then 0.025
			else 10000 end
		when ficoScore>=660 then 
			case
			when LoanToValue<=.6 then .0025
			when LoanToValue<=.75 then 0.0075
			when LoanToValue<=.80 then 0.015
			when LoanToValue<=.85 then 0.025 
			else 10000 end
		when ficoScore>=640 then
			case
			when LoanToValue<=.6 then .0025
			when LoanToValue<=.75 then 0.0125
			when LoanToValue<=.80 then 0.0225
			when LoanToValue<=.85 then 0.03
			else 10000 end
		when ficoScore>=620 then 
			case
			when LoanToValue<=.6 then .0025
			when LoanToValue<=.75 then 0.0125
			when LoanToValue<=.80 then 0.0275
			when LoanToValue<=.85 then 0.03 
			else 10000 end
		else  
			case
			when LoanToValue<=.6 then .0125
			when LoanToValue<=.75 then 0.0225
			when LoanToValue<=.80 then 0.0275
			when LoanToValue<=.85 then 0.03 
			else 10000 end
		end
	else 0 end AS px	
FROM MPFLoan


---- uv_FNpxSubdebt
alter  view uv_FNpxSubdebt as 
SELECT LoanNumber,  TotalLoanToValue,  LoanToValue, ficoScore, 
	case 
	when TotalLoanToValue<>0 AND TotalLoanToValue<>LoanToValue then
		case 
		when LoanToValue<=.65 then
			case 
			when TotalLoanToValue<=.8 then 0
			when TotalLoanToValue<=.95 then
				case when ficoScore<720 then 0.005 else 0.0025 end
			else 10000 end 
		when LoanToValue<=.75 then
			case
			when TotalLoanToValue<=.8 then 0
			when TotalLoanToValue<=.95 then
				case when ficoScore<720 then 0.0075 else 0.005 end
			else 10000 end 
		when LoanToValue<=.95 then
			case
			when TotalLoanToValue<=.75 then 0
			when TotalLoanToValue<=.95 then 
				case when ficoScore<720 then 0.01 else 0.0075 end
			when TotalLoanToValue<=.97 then .015
			else 10000 end 
		else 10000 end
	else 0 end as px
FROM MPFLoan


---- uv_FNpxPropertyType
alter view uv_FNpxPropertyType as 
SELECT  loanNumber, PropertyType, ficoScore, loanToValue,
	case 
	when PropertyType in ('PT04','PT05','PT08') then 
		case when LoanToValue<=.85 then 0.01 else 10000 end
	when PropertyType='PT11' then 
		case when LoanToValue<=.95 then 0.005 else 10000 end
	when PropertyType in ('PT06','PT07','PT15','PT16','PT17','PT18','PT19') then
		case when LoanToValue<=.75 then 0 when (LoanToValue>.75 AND LoanToValue<=.97) then 0.0075 else 10000 end 
	when PropertyType in ('PT09','PT10') then 
		case when LoanToValue<=.75 then 0.01 else 10000 end
	else 0 
	end AS px
FROM MPFLoan


--- uv_FNpxPMI
alter view uv_FNpxPMI as 
SELECT LoanNumber,  TotalLoanToValue,  LoanToValue, ficoScore, PMIPercent,
	case when PMIPercent>0 then
		case 
		when LoanToValue>.97 then  10000
		when LoanToValue>.95 then
			case
			when ficoScore>=740 then .01
			when ficoScore>=700 then .0125
			when ficoScore>=680 then .0175 
			when ficoScore>=660 then .02125
			when ficoScore>=640 then .02375
			when ficoScore>=620 then .0275
			else .03 end
		when LoanToValue>.90 then
			case 
			when ficoScore>=740 then .005
			when ficoScore>=680 then .00875
			when ficoScore>=660 then .0175
			when ficoScore>=640 then .02
			when ficoScore>=620 then .0225
			else .025 end
		when LoanToValue>.85 AND numberOfMonths>240 then
			case
			when ficoScore>=740 then .00375 
			when ficoScore>=720 then .00625
			when ficoScore>=680 then .0075
			when ficoScore>=660 then .0125
			when ficoScore>=640 then .0175
			when ficoScore>=620 then .02 
			else .0225 end
		when LoanToValue>.8 AND numberOfMonths>240 then
			case
			when ficoScore>=680 then .00125
			when ficoScore>=660 then .0075
			when ficoScore>=640 then .0125
			when ficoScore>=620 then .0175
			else .02 end 
		else 0 end
	else 0 end as px
FROM MPFLoan

--- uv_FNpxInvestment -------- 20140508 count second home as investment
create view uv_FNpxInvestment as 
SELECT  loanNumber, loanToValue, Occupancy,
	case 
	when Occupancy LIKE 'P%' then 0.0
	else 
		case 
		when LoanToValue<=.75 then 0.0175 
		when LoanToValue<=.8 then 0.03
		when LoanToValue<=.85 then 0.0375
		else 10000.0 end
	end AS px
FROM MPFLoan


--================================================================================================
--- uv_MPFLLPA ---> MPFLLPA ----- 20140508 added Occupancy and invPx
create view uv_MPFLLPA as
SELECT f.loanNumber, f.FICOScore, f.LoanToValue, lp.LoanPurpose, pmi.TotalLoanToValue, pt.PropertyType, inv.occupancy, 100*(f.px+lp.px+pmi.px+pt.px+sd.px+inv.px) as LLPA, 
	100*f.px as fPx, 100*lp.px as lpPx, 100*pmi.px as pmiPx, 100*pt.px as ptPx, 100*sd.px as sdPx, 100*inv.px as invPx, curBal, PFInumber, origYear
FROM ((((uv_FNPxFICOLTV as f INNER JOIN uv_FNPxCashout as lp ON f.loanNumber = lp.loanNumber) 
INNER JOIN uv_FNPxPMI as pmi ON lp.loanNumber = pmi.LoanNumber) 
INNER JOIN uv_FNPxPropertyType as pt ON pmi.LoanNumber = pt.loanNumber) 
INNER JOIN uv_FNPxSubDebt as sd ON pt.loanNumber = sd.LoanNumber)
INNER JOIN uv_FNPxInvestment as inv ON f.loanNumber = inv.LoanNumber


CREATE TABLE [dbo].[MPFLLPA](
	[loanNumber] [int] NOT NULL,
	[FICOScore] [real] NULL,
	[LoanToValue] [decimal](7, 6) NULL,
	[LoanPurpose] [varchar](4) NULL,
	[TotalLoanToValue] [decimal](7, 6) NULL,
	[PropertyType] [varchar](4) NULL,
	[Occupancy] [varchar](25) NULL,
      fPx   real NULL,
      lpPx  real NULL,
      pmiPx real NULL,
      ptPx  real NULL,
      sdPx  real NULL,
      invPx real NULL,
	[LLPA] [numeric](19, 5) NULL,
	[curBal] [money] NULL,
	[PFInumber] [int] NULL,
	[PFIName] [varchar](50) NULL,
	[origYear] [decimal](4, 0) NULL,
	origMonth CHAR(3) NULL,
	[PFIOwner] [varchar](100) NULL,
	[ruleSetID] [int] NULL,
	[ProgramCode] [varchar](6) NULL
) ON [PRIMARY]


--------- load MPFLLPA
INSERT INTO MPFLLPA (
	  [loanNumber]
      ,[FICOScore]
      ,[LoanToValue]
      ,[LoanPurpose]
      ,[TotalLoanToValue]
      ,[PropertyType]
      ,Occupancy
      ,fPx
      ,lpPx
      ,pmiPx
      ,ptPx
      ,sdPx
      ,invPx
      ,LLPA
      ,[curBal]
      ,[PFInumber]
      ,[PFIName]
      ,[origYear]
	  ,origMonth
      ,PFIOwner
      ,ruleSetID
      ,[ProgramCode])
SELECT l.[loanNumber]
      ,l.[FICOScore]
      ,l.[LoanToValue]
      ,l.[LoanPurpose]
      ,l.[TotalLoanToValue]
      ,l.[PropertyType]
      ,l.Occupancy
      ,v.fPx
      ,v.lpPx
      ,v.pmiPx
      ,v.ptPx
      ,v.sdPx
      ,v.invPx
      ,v.LLPA
      ,l.[curBal]
      ,l.[PFInumber]
      ,l.[PFIName]
      ,l.[origYear]
	  ,convert (char(3), l.fundingDate, 9) as origMonth
      ,l.PFIOwner
      ,m.ruleSetID
      ,l.[ProgramCode]
FROM ([uv_MPFLLPA] as v inner join uv_MPFLoan as l on v.loanNumber=l.loanNumber)
inner join mraTest.dbo.MasterCommitment_NOSHFD as m on l.MANumber=m.MANumber


