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

	