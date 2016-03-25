Option Compare Database
Option Explicit

Private Sub Command23_Click()
Dim stDocName As String
    Dim stLinkCriteria As String
    
    DoCmd.Close acForm, "frmStartUp"

    stDocName = "frmNetCouponAdjustmentFactors"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
End Sub

Private Sub cmdCheckHFSBalance_Click()
    cmdCheckHFSBalance.ForeColor = 8421504
    DoCmd.SetWarnings False
    DoCmd.OpenReport "rptHFSEODCheck", acViewReport
    
    DoCmd.SetWarnings True
End Sub

Private Sub cmdCheckMissingHFSPrices_Click()
    cmdCheckMissingHFSPrices.ForeColor = 8421504
    
    DoCmd.SetWarnings False
    '   Show loans in HFSLoan Table without price
    DoCmd.OpenQuery "qryMissingHFSPrices", acNormal, acEdit
    
    MsgBox "Process Complete", vbInformation, " "
    
    DoCmd.SetWarnings True
End Sub

Private Sub cmdDeleteMiddleware_Click()

cmdDeleteMiddleware.ForeColor = 8421504

DoCmd.SetWarnings False
DoCmd.OpenQuery "qryDeleteMPFandForwardsfromInstrumenttblMiddleware", acNormal, acEdit
'Added ???
DoCmd.OpenQuery "qryDeleteDataFromPortfolioContentsMiddleware", acNormal, acEdit
 MsgBox "Process Complete", vbInformation, " "
    
    DoCmd.SetWarnings True


End Sub

Private Sub cmdExportHFSToTPODS_Click()
    cmdExportHFSToTPODS.ForeColor = 8421504
    Dim rs As DAO.Recordset, QDF As DAO.QueryDef, P As Variant, col As Variant, f As Variant
    Dim F1 As Integer, separator As String, fields As String, FileName As String, currYr As String, currMo As String, currDay As String, TPODSLoc As String
    Dim fso As FileSystemObject
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    

    
    
    currYr = Year(Date)
    currMo = Format(Date, "mm")
    currDay = Format(Date, "dd")
    FileName = "K:\MRO\Archive\" & currYr & "\" & currYr & "-" & currMo & "\" & currYr & "-" & currMo & "-" & currDay & "\Pricing\Pricing-HFV-Loans.csv"
    ' TPODSLoc = "\\prodfs\SOFTWAREp\BANKAPPS\TPO\Poly\In\"
    ' TPODSLoc = "\\testfs\SOFTWARET\DP\Haus\Data\"
	
    
    ' FileName = "S:\share\msharan\2014\2014-04\GNMA\Pricing-HFV-Loans.csv"
    
        
    'Open query
    Set QDF = CurrentDb().QueryDefs("qryHFSLoanPricesToTPODS")
    Set rs = QDF.OpenRecordset
    
    'Open the file for output
    F1 = 1
    Open FileName For Output As F1
    
    'Loop over field names and output them to the file
    separator = ""
    fields = ""
    For Each col In rs.fields
        fields = fields & separator & col.Name
        separator = ","
    Next col
    Print #F1, fields
    'MsgBox fields
    
    'loop over query and output all fields
    fields = ""
    While Not rs.EOF
        fields = ""
        separator = ""
        For Each f In rs.fields
            fields = fields & separator & f.Value
            separator = ","
        Next f
        Print #F1, fields
    rs.MoveNext
    Wend
    
    'close file and query
    Close #F1
    QDF.Close
    rs.Close
    Set rs = Nothing
    Set QDF = Nothing
    
    ' Hua 20150209 No access yet. fso.CopyFile FileName, TPODSLoc, True
End Sub

Private Sub cmdInsertMiddleware_Click()

cmdInsertMiddleware.ForeColor = 8421504

DoCmd.SetWarnings False
    '   Append data to Instrument Table
    DoCmd.OpenQuery "qryAppendMPFSecuritiesToInstrumentTableMiddleware", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFMarkToModelToInstrumentTableMiddleware", acNormal, acEdit
    '   Append data to Portfolio Contents
    DoCmd.OpenQuery "qryAppendMPFSecuritiesToPortfolioContentsTableMiddleware", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFMarkToModelToPortfolioContentsTableMiddleware", acNormal, acEdit
    '   Append data to Instrument Table
    DoCmd.OpenQuery "qryAppendMPFForwardCommitmentsToInstrumentMiddleware", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFForwardMarkToModelToInstrumentMiddleware", acNormal, acEdit
    '   Append data to Portfolio Contents
    DoCmd.OpenQuery "qryAppendMPFForwardCommitmentsToPortfolioContentsMiddleware", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFForwardMarkToModelToPortfolioContentsMiddleware", acNormal, acEdit
        
    DoCmd.OpenQuery "qryUpdateMiddlewareLoadStats", acViewNormal, acEdit
        
    MsgBox "Process Complete", vbInformation, " "
    
    DoCmd.SetWarnings True
End Sub

Private Sub cmdLoadHFSLoansEOD_Click()
    cmdLoadHFSLoansEOD.ForeColor = 8421504
    
    DoCmd.SetWarnings False
    '   Append data to HFSLoan Table
    DoCmd.OpenQuery "qryInsertHFSLoansEOD", acNormal, acEdit
    DoCmd.OpenQuery "qryDeleteClosedHFSLoans", acNormal, acEdit
    
    MsgBox "Process Complete", vbInformation, " "
    
    DoCmd.SetWarnings True
End Sub

Private Sub cmdShowMissingForwards_Click()

cmdShowMissingForwards.ForeColor = 8421504

    DoCmd.OpenReport "rptMPFForwardsMissingPrice", acViewPreview

End Sub

Private Sub cmdShowMPFCusips_Click()

On Error GoTo err_handler

cmdShowMPFCusips.ForeColor = 8421504

DoCmd.OpenQuery "qryShowMPFCUSIPS", acViewNormal, acReadOnly

Exit Sub

err_handler:

MsgBox Err.Description, vbCritical, "Error Occured"

End Sub

Private Sub cmdUpdateHFSLoanPrices_Click()
    cmdUpdateMPFPrice.ForeColor = 8421504
    
    DoCmd.SetWarnings False
    ' Update Price for HFS Loans Using MPFPrice.XLS
    DoCmd.OpenQuery "qryUpdateHFSLoanPrice", acNormal, acEdit
    
    MsgBox "Process Complete", vbInformation, " "
    DoCmd.SetWarnings True
End Sub


'========== Update settled MPF Price
Private Sub cmdUpdateMPFPrice_Click()
    
    cmdUpdateMPFPrice.ForeColor = 8421504
    
    DoCmd.SetWarnings False
    ' Update Price for MPF Loans Using MPFPrice.XLS
    DoCmd.OpenQuery "qryUpdatePricesIntblSecuritiesPalmsSourceTest", acNormal, acEdit
    DoCmd.OpenQuery "qryUpdatePricesIntblSecuritiesPalmsTestSourceMarkToModel", acNormal, acEdit
    MsgBox "Process Complete", vbInformation, " "
    DoCmd.SetWarnings True
    
End Sub

'========== load loan level data
Private Sub Command24_Click()
On Error GoTo Command24_Err

    Command24.ForeColor = 8421504

    DoCmd.SetWarnings False
    
    'Update tblPalmsrRepo
    'DoCmd.OpenQuery "qryUpdateDailyRepoRate", acNormal, acEdit 'UNUSED and REMOVED M.Butler - 11/21/2006
    
    'Delete Forward Settle DCs from table MPFHedges
    DoCmd.OpenQuery "DeleteForwardSettleDCs MPFHedges", acNormal, acEdit
    
    'Make table with Master Agreement data used
    DoCmd.OpenQuery "qryMaketbltempMA", acNormal, acEdit
    
    
    'Create table tblLoanFundingDateViewTABLE
   ' DoCmd.OpenQuery "qryDeletetblLoanFundingDateViewtoTABLE", acNormal, acEdit
   ' DoCmd.OpenQuery "qryAppendtblLoanFundingDateViewtoTABLE", acNormal, acEdit
    
    
    'Delete data in tblMPFLoansAggregateSource
    DoCmd.OpenQuery "qryDeletetblMPFLoansAggregateSource", acNormal, acEdit
    DoCmd.OpenQuery "qryDeletetbltempMPF", acNormal, acEdit
    
    'Append new data to tblMPFLoansAggregateSource
    DoCmd.OpenQuery "qryAppendAllLoans", acNormal, acEdit
    
    'Assign Action Code.
    DoCmd.OpenQuery "qryCreateTempActionCode", acNormal, acEdit
    DoCmd.OpenQuery "qryUpdateLoanActionCode", acNormal, acEdit
    
    'Delete Closed Loans
    DoCmd.OpenQuery "qryDeleteClosedLoans", acNormal, acEdit
    
    
    DoCmd.OpenQuery "qryAppendAllLoansStepTwo", acNormal, acEdit
    'DoCmd.OpenQuery "qryDeletetbltempMPF", acNormal, acEdit

    'Delete loans in tblMPFLoansAggregateSource where Chicago Participation is 0%.
    DoCmd.OpenQuery "qryDeleteChicagoParticipation0", acNormal, acEdit
 
    
    
 MsgBox "Loan Level Data Loaded", vbInformation, " "
        
    
DoCmd.SetWarnings True
   

Command24_Exit:
    Exit Sub

Command24_Err:
    MsgBox Error$
    Resume Command24_Exit

End Sub

Private Sub Command25_Click()
Dim stDocName As String
    Dim stLinkCriteria As String
    
DoCmd.Close acForm, "frmStartUp"

stDocName = "frmMissingEOMRatesandFees"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

End Sub

Private Sub Command26_Click()
On Error GoTo Err_Command26_Click

    Command26.ForeColor = 8421504
   
    Dim stDocName As String
    
    'DoCmd.Close acForm, "frmStartUp"
    
    stDocName = "rptProgramBalance"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command26_Click:
    
    Exit Sub

Err_Command26_Click:
    MsgBox Err.Description
    Resume Exit_Command26_Click
End Sub

Private Sub Command27_Click()
On Error GoTo Err_Command27_Click

    Dim stDocName As String

    
    stDocName = "rptMPFLoansForPALMS"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command27_Click:
    Exit Sub

Err_Command27_Click:
    MsgBox Err.Description
    Resume Exit_Command27_Click
End Sub

Private Sub Command29_Click()
    DoCmd.SetWarnings False
    '   Append data to Instrument Table
    DoCmd.OpenQuery "qryAppendMPFSecuritiesToInstrumentTable", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFForwardCommitmentsToInstrument", acNormal, acEdit
    '   Append data to Portfolio Contents
    DoCmd.OpenQuery "qryAppendMPFSecuritiesToPortfolioContentsTable", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFForwardCommitmentsToPortfolioContents", acNormal, acEdit
    
    
    MsgBox "Process Complete", vbInformation, " "
    
    DoCmd.SetWarnings True
End Sub

Private Sub Command32_Click()

DoCmd.SetWarnings False


' Update OAS for NonSeasoned MPF Loans
DoCmd.OpenQuery "qryUpdateTodaysSpreadsNonseasoned", acNormal, acEdit
' Update New Seasoned Loan Prices
DoCmd.OpenQuery "qryUpdatePricesIntblSecuritiesPalmsSource", acNormal, acEdit
MsgBox "Process Complete", vbInformation, " "
DoCmd.SetWarnings True

End Sub

Private Sub Command35_Click()
On Error GoTo Err_Command35_Click

    Dim stDocName As String

    stDocName = "rptMPFPortfolioSummary"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command35_Click:
    Exit Sub

Err_Command35_Click:
    MsgBox Err.Description
    Resume Exit_Command35_Click
End Sub

Private Sub Command36_Click()


Dim stDocName As String
Dim stLinkCriteria As String
    
DoCmd.Close acForm, "frmStartUp"

stDocName = "frmUpdateFowardPrices"
DoCmd.OpenForm stDocName, , , stLinkCriteria



End Sub

Private Sub Command38_Click()
On Error GoTo Err_Command38_Click

    Command38.ForeColor = 8421504
    
    Dim stDocName As String
    
    'DoCmd.Close acForm, "frmStartUp"
    
    stDocName = "rptMPFForwards"
    DoCmd.OpenReport stDocName, acPreview


Exit_Command38_Click:
    Exit Sub

Err_Command38_Click:
    MsgBox Err.Description
    Resume Exit_Command38_Click
    
End Sub


'========= Aggregate MPF Forwards
Private Sub Command39_Click()
 
Command39.ForeColor = 8421504

 DoCmd.SetWarnings False

 'Delete
 DoCmd.OpenQuery "qryDeleteDataFromTblForwardCommitmentPalmsSource", acNormal, acEdit
 'Delete
 DoCmd.OpenQuery "qryDeleteDataFromTblForwardCommitmentMarkToModelLoans", acNormal, acEdit
 
 
 DoCmd.OpenQuery "qryAggregateForwardSettleCommitments", acNormal, acEdit
 
 'As of 2/15/05, this will not be needed.  Hedged loans are identified at a loan level
 'and forward commitments are hedged using another process.  Greg Fleming
   'DoCmd.OpenQuery "qryAggregateHedgedForwardSettleCommitments", acNormal, acEdit
   'DoCmd.OpenQuery "qryAggregateHedgedForwardSettleMarkToModelCommitments", acNormal, acEdit
 
 DoCmd.OpenQuery "AppendForwardSettleDCs MPFHedges", acNormal, acEdit
 
 DoCmd.OpenQuery "UpdateMPFForward15&20PPMult", acNormal, acEdit
 DoCmd.OpenQuery "UpdateMPFForward30PPMult", acNormal, acEdit
 
 DoCmd.OpenQuery "UpdateMPFForward15&20PPMultMarkToModel", acNormal, acEdit
 DoCmd.OpenQuery "UpdateMPFForward30PPMultMarkToModel", acNormal, acEdit
 

 MsgBox "Process Complete", vbInformation, " "
    DoCmd.SetWarnings True
End Sub



'========== Run Palms data report
Private Sub Command40_Click()

On Error GoTo Err_Command40_Click

    Command40.ForeColor = 8421504

    Dim stDocName As String

    DoCmd.SetWarnings False
    
    '  This will change product coding in Loan Table to the MS - for FlexiSwap Code
    DoCmd.OpenQuery "qryUpdateFlexiSwap", acNormal, acEdit
    
    ' This will change remittance type for nine seasoned loans from AA to SS
    DoCmd.OpenQuery "qryUpdateRemitTypeNorthFederalSavingsLoans", acNormal, acEdit
    
    '    Delete data from tblSecuritiesPalmsSource
    DoCmd.OpenQuery "qryDeleteDataFromtblSecuritiesPalmsSource", acNormal, acEdit
    
    'Mark To Model
    DoCmd.OpenQuery "qryDeleteDataFromtblMarkToModelLoans", acNormal, acEdit
    
    '    Append new data to tblSecuritiesPalmsSource - Standard
    DoCmd.OpenQuery "qryAppendUnhedgedDataTotblSecuritiesPalmsSource", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendUnhedgedDataTotblSecuritiesPalmsSource - NONSEASONED", acNormal, acEdit
    
    '    Append new data to Mark To Model
    'DoCmd.OpenQuery "qryAppendUnhedgedDataTotblMarkToModelLoans", acNormal, acEdit
    'DoCmd.OpenQuery "qryAppendUnhedgedDataTotblMarkToModelLoans - NONSEASONED", acNormal, acEdit
    
    'Standard
    DoCmd.OpenQuery "qryAppendHedgedDataTotblSecuritiesPalmsSource", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendHedgedDataTotblSecuritiesPalmsSource - NONSEASONED", acNormal, acEdit
    
    'Mark To Model
    DoCmd.OpenQuery "qryAppendhedgedDataTotblMarkToModelLoans", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendhedgedDataTotblMarkToModelLoans - NONSEASONED", acNormal, acEdit

    'Standard
    DoCmd.OpenQuery "UpdateMPF15&20PPMult", acNormal, acEdit
    DoCmd.OpenQuery "UpdateMPF30PPMult", acNormal, acEdit
    
    'Mark To Model
    DoCmd.OpenQuery "UpdateMPF15&20PPMultMarkToModel", acNormal, acEdit
    DoCmd.OpenQuery "UpdateMPF30PPMultMarkToModel", acNormal, acEdit
    
    
    DoCmd.SetWarnings True

    stDocName = "rptMPFLoansForPALMS"
    DoCmd.OpenReport stDocName, acPreview
    stDocName = "rptMPFLoansForPalmsMarkToModel"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command40_Click:
    Exit Sub

Err_Command40_Click:
    MsgBox Err.Description
    Resume Next
    Resume Exit_Command40_Click
End Sub





Private Sub Command42_Click()
DoCmd.SetWarnings False
    '   Append data to Instrument Table
    DoCmd.OpenQuery "qryAppendMPFSecuritiesToInstrumentTableGAP", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFForwardCommitmentsToInstrumentGAP", acNormal, acEdit
    '   Append data to Portfolio Contents
    DoCmd.OpenQuery "qryAppendMPFSecuritiesToPortfolioContentsTableGAP", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFForwardCommitmentsToPortfolioContentsGAP", acNormal, acEdit
    
    
    MsgBox "Process Complete", vbInformation, " "
    
    DoCmd.SetWarnings True

End Sub

Private Sub Command43_Click()
DoCmd.SetWarnings False
' Delete Old Data
DoCmd.OpenQuery "qryDeleteEOMRatesAndFees"

' Append New data from Excel Sheet
DoCmd.OpenQuery "qryAppendEOMSchedulestoTable"
 
MsgBox "Schedule Update Complete", vbInformation, " "

DoCmd.SetWarnings True

End Sub

'======= MPF Loans Open report
Private Sub Command44_Click()

On Error GoTo Err_Command44_Click

    Command44.ForeColor = 8421504
    
    Dim stDocName As String

        stDocName = "rptMPFLoansForPALMS"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command44_Click:
    Exit Sub

Err_Command44_Click:
    MsgBox Err.Description
    Resume Exit_Command44_Click

End Sub

Private Sub Command45_Click()
DoCmd.SetWarnings False
    '   Append data to Instrument Table
    DoCmd.OpenQuery "qryAppendMPFSecuritiesToInstrumentTable - Playspace", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFForwardCommitmentsToInstrument - Playspace", acNormal, acEdit
    '   Append data to Portfolio Contents
    DoCmd.OpenQuery "qryAppendMPFSecuritiesToPortfolioContentsTable - Playspace", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFForwardCommitmentsToPortfolioContents - Playspace", acNormal, acEdit
    
    
    MsgBox "Process Complete", vbInformation, " "
    
    DoCmd.SetWarnings True
End Sub

Private Sub Command46_Click()

DoCmd.SetWarnings False
'    Update OAS for Pass Thru Rate = 7
DoCmd.OpenQuery "qryUpdateMPFOASforPTRate7", acNormal, acEdit
   MsgBox "Process Complete", vbInformation, " "
    
    DoCmd.SetWarnings True
End Sub

Private Sub Command47_Click()
On Error GoTo Err_Command47_Click

    Dim stDocName As String

        stDocName = "rptMPFLoansForPALMS"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command47_Click:
    Exit Sub

Err_Command47_Click:
    MsgBox Err.Description
    Resume Exit_Command47_Click
End Sub

Private Sub Command48_Click()
DoCmd.SetWarnings False
DoCmd.OpenQuery "qryUpdateForwardCommitmentOAS", acNormal, acEdit
DoCmd.OpenQuery "qryUpdateForwardCommitmentOASMarkToModel", acNormal, acEdit
 MsgBox "Process Complete", vbInformation, " "
    
    DoCmd.SetWarnings True
End Sub

Private Sub Command49_Click()
DoCmd.SetWarnings False
    '   Append data to Instrument Table
    DoCmd.OpenQuery "qryAppendMPFSecuritiesToInstrumentTable-QPalms1-2", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFForwardCommitmentsToInstrument-QPalms1-2", acNormal, acEdit
    '   Append data to Portfolio Contents
    DoCmd.OpenQuery "qryAppendMPFSecuritiesToPortfolioContentsTable-QPalms1-2", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFForwardCommitmentsToPortfolioContents-QPalms1-2", acNormal, acEdit
    
    MsgBox "Process Complete", vbInformation, " "
    
    DoCmd.SetWarnings True
End Sub



'========== update missing dates
Private Sub Command50_Click()

Command50.ForeColor = 8421504

DoCmd.SetWarnings False
    ' Add Date for Missing Original Loan Closing Date and for Dates 1/1/1900
    DoCmd.OpenQuery "qryUpdateOriginalLoanClosingDateIntblMPFLoanAggregateSource", acNormal, acEdit
    DoCmd.OpenQuery "qryUpdateOriginalLoanClosingDate(MissingData)", acNormal, acEdit
    DoCmd.OpenQuery "qryUpdateProductCodeforFX20bookedthirty", acNormal, acEdit
    'DoCmd.OpenQuery "qryUpdateProductCodeforFX15bookedthirty", acNormal, acEdit
    MsgBox "Process Complete", vbInformation, " "
    
    DoCmd.SetWarnings True

End Sub

'========== Export data to Palm Production (not used)
Private Sub Command52_Click()
DoCmd.SetWarnings False
    '   Append data to Instrument Table
    DoCmd.OpenQuery "qryAppendMPFSecuritiesToInstrumentTableP", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFMarkToModelToInstrumentTableP", acNormal, acEdit

    '   Append data to Portfolio Contents
    DoCmd.OpenQuery "qryAppendMPFSecuritiesToPortfolioContentsTableP", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFMarkToModelToPortfolioContentsTableP", acNormal, acEdit

    '   Append data to Instrument Table
    DoCmd.OpenQuery "qryAppendMPFForwardCommitmentsToInstrumentP", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFForwardMarkToModelToInstrumentP", acNormal, acEdit

    '   Append data to Portfolio Contents
    DoCmd.OpenQuery "qryAppendMPFForwardCommitmentsToPortfolioContentsP", acNormal, acEdit
    DoCmd.OpenQuery "qryAppendMPFForwardMarkToModelToPortfolioContentsP", acNormal, acEdit

    MsgBox "Process Complete", vbInformation, " "

    DoCmd.SetWarnings True

End Sub

'============= Delete data from Palms Production -- Not used
Private Sub Command59_Click()
DoCmd.SetWarnings False
DoCmd.OpenQuery "qryDeleteMPFandForwardsfromInstrumenttblProduction", acNormal, acEdit
 MsgBox "Process Complete", vbInformation, " "
    
    DoCmd.SetWarnings True

End Sub

Private Sub Command71_Click()
Dim stDocName As String

DoCmd.SetWarnings False
 DoCmd.OpenQuery "qryDeleteDataFromTblForwardCommitmentPalmsSourcedaily", acNormal, acEdit
 DoCmd.OpenQuery "qryAggregateForwardSettleCommitmentsdaily", acNormal, acEdit
 'Don't run this query daily because pulling in prices from TPM
 'DoCmd.OpenQuery "qryUpdateForwardCommitmentOASdaily", acNormal, acEdit
 DoCmd.Close acForm, "frmStartUp"
    
    stDocName = "rptMPFForwardsdaily"
    DoCmd.OpenReport stDocName, acPreview


Exit_Command71_Click:
    Exit Sub

Err_Command71_Click:
    MsgBox Err.Description
    Resume Exit_Command71_Click
    
 
End Sub

Private Sub Command72_Click()

DoCmd.SetWarnings False
 DoCmd.OpenQuery "qryDeleteMPFDCfromPlayspacedaily", acNormal, acEdit
 DoCmd.OpenQuery "qryAppendMPFFwdstoInstrument-Playdaily", acNormal, acEdit
 DoCmd.OpenQuery "qryAppendMPFFwdtoPortContPlaydaily", acNormal, acEdit
 
MsgBox "Process Complete", vbInformation, " "
End Sub

Private Sub Command73_Click()

On Error GoTo err_handler

    Dim sTime, dTime, fTime, bTime, eTime, sStep
    

    
    DoCmd.SetWarnings False
    
    'Update tblPalmsrRepo
    sStep = "Updating Repo Rate"
    DoCmd.OpenQuery "qryUpdateDailyRepoRate", acNormal, acEdit
    
    'Delete Forward Settle DCs from table MPFHedges
    dTime = Now()
    sStep = "Delete Forwards"
    DoCmd.OpenQuery "DeleteForwardSettleDCs MPFHedges", acNormal, acEdit
    
    'Delete data in tblMPFLoansAggregateSource
    sStep = "Clear Aggregate Table"
    DoCmd.OpenQuery "qryDeletetblMPFLoansAggregateSource", acNormal, acEdit
    
    'Append new data to tblMPFLoansAggregateSource -- this is what the Pass-Through would replace
    fTime = Now()
    sStep = "Append Flow Loans"
    DoCmd.OpenQuery "qryFlowToAggregateUsingPassthroughSQL", acNormal, acEdit
    bTime = Now()
    sStep = "Append Batch Loans"
    DoCmd.OpenQuery "qryBatchtoAggregateUsingPassThroughSQL", acNormal, acEdit
    
    eTime = Now()
    
    DoCmd.SetWarnings True
    
    MsgBox "Loan Level Data Loaded" & vbCrLf & _
            "Started Deletions: " & dTime & DateDiff("s", dTime, fTime) & " seconds)" & vbCrLf & _
            "Started Flow Loans: " & dTime & DateDiff("n", fTime, bTime) & " minutes)" & vbCrLf & _
            "Started Batch Loans: " & dTime & DateDiff("n", bTime, eTime) & " minutes)" & vbCrLf & _
            "Completed: " & eTime & vbCrLf & _
            "Total Execution Time: " & DateDiff("n", sTime, eTime), vbInformation, "Load Completion Statistics"
    
    Exit Sub

err_handler:

    eTime = Now()
    MsgBox "Error Number: " & Err.Number & " '" & Err.Description & "' was encountered during step '" & _
                sStep & "'" & vbCrLf & "The Process had run " & _
                DateDiff("n", sTime, eTime) & " minutes.", vbCritical, "Contact Application Support"

End Sub

'============= clear tables before compact
Private Sub Command76_Click()

Command76.ForeColor = 8421504

DoCmd.SetWarnings False

    DoCmd.OpenQuery "qryDeletetblLoanFundingDateViewtoTABLE", acNormal, acEdit
    DoCmd.OpenQuery "qryDeletetblMPFLoansAggregateSource", acNormal, acEdit
    DoCmd.OpenQuery "qryDeletetbltempMPF", acNormal, acEdit

DoCmd.SetWarnings True

End Sub


'========== Collect MPF forward level data
Private Sub Command78_Click()
 
Command78.ForeColor = 8421504
 
 DoCmd.SetWarnings False
 'DoCmd.OpenQuery "qryFundedandUnfundedDCsbyPFI", acNormal, acEdit
 DoCmd.OpenQuery "qryUpdateMPFForwardPrices", acNormal, acEdit
 DoCmd.OpenForm "frmShowMissingMPFForwardPrice", , , , , acDialog
 
 MsgBox "Process Complete", vbInformation, " "
    DoCmd.SetWarnings True
    
End Sub


'========== load Chicago DC's
Private Sub Command80_Click()
Command80.ForeColor = 8421504

 DoCmd.SetWarnings False
 
 'DoCmd.OpenQuery "QryDeliveryCommitment_ChicagoOnly", acNormal, acEdit
 DoCmd.OpenQuery "qryFundedandUnfundedDCsbyPFI", acNormal, acEdit
 
 MsgBox "Load Chicago DC's Process Complete", vbInformation, " "
    DoCmd.SetWarnings True

End Sub

Private Sub ExitButton_Click()
DoCmd.Close

Exit_cmdClose_Click:
    Exit Sub

Err_cmdClose_Click:
    MsgBox Err.Description
    Resume Exit_cmdClose_Click
End Sub

Private Sub SetEOMSchedule_Click()
On Error GoTo Err_SetEOMSchedule_Click

        
    SetEOMSchedule.ForeColor = 8421504
        
    Dim stDocName As String
    Dim stLinkCriteria As String
    
    DoCmd.Close acForm, "frmStartUp"

    stDocName = "frmSetEOMSchedule"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_SetEOMSchedule_Click:
    Exit Sub

Err_SetEOMSchedule_Click:
    MsgBox Err.Description
    Resume Exit_SetEOMSchedule_Click
    
End Sub

Private Sub Command28_Click()
On Error GoTo Err_Command28_Click

    Dim stDocName As String

    stDocName = "frmPrepaymentFactors"
    DoCmd.OpenForm stDocName

    DoCmd.Close acForm, "frmStartUp"

Exit_Command28_Click:
    Exit Sub

Err_Command28_Click:
    MsgBox Err.Description
    Resume Exit_Command28_Click
    
End Sub
