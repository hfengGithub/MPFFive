'========================== basAgentFee
Option Compare Database
Option Explicit

'Purpose:   To calculate an interpolated agent fee given an interest rate
'           and a schedule code.
'Parameters:
'  dblInterestRate   in - Interest rate for the loan
'  sScheduleCode     in - Code for the schedule of rates and fees
'  return value         - Interpolated agent fee as a currency.
'
'History:
'  09-26-97 Zeke Leventhal-Arnold
'     Created to allow anticipated agent fee to be calculated from Loan
'     Funding screen as well as New Loan Funding screen.
'
Public Function curAgentFee(ByVal dblInterestRate As Double, _
      ByVal sScheduleCode As String) As Double
      
      
On Error GoTo curAgentFee_Error
   
   Dim dbCurrent As Database
   Dim rsUpperRate As Recordset
   Dim rsLowerRate As Recordset
   Dim dblUpperRate As Double
   Dim dblLowerRate As Double
   Dim dblUpperFee As Double
   Dim dblLowerFee As Double
   Dim strSQL As String
   Dim bCancel As Boolean
   
   ' Set a default agent fee amount to be returned in case of any errors.
   curAgentFee = 0
   
   Set dbCurrent = CurrentDb
   ' Get upper rate
   strSQL = "Select Top 1 * from tblLoanRatesAgentFees where ScheduleCode = '" & _
         sScheduleCode & "' And Rate >= " & dblInterestRate & " Order By Rate Asc"
   Set rsUpperRate = dbCurrent.OpenRecordset(strSQL, dbOpenSnapshot)
   If Not rsUpperRate.EOF Then
      dblUpperRate = rsUpperRate!Rate
      dblUpperFee = rsUpperRate!Fee
   Else
      Beep
      MsgBox "Interest Rate is not valid for this schedule."
      bCancel = True
   End If
   
   If Not bCancel Then
      'Get lower rate
      strSQL = "Select Top 1 * from tblLoanRatesAgentFees where ScheduleCode = '" & _
            sScheduleCode & "' And Rate <= " & dblInterestRate & " Order By Rate Desc"
      Set rsLowerRate = dbCurrent.OpenRecordset(strSQL, dbOpenSnapshot)
      If Not rsLowerRate.EOF Then
         dblLowerRate = rsLowerRate!Rate
         dblLowerFee = rsLowerRate!Fee
      Else
         Beep
         MsgBox "Interest Rate is not valid for this schedule."
         bCancel = True
      End If
   End If
   
   If Not bCancel Then
      ' Calculate interpolated agent fee rate
      curAgentFee = Interpolate(dblInterestRate, dblLowerRate, dblUpperRate, _
            dblLowerFee, dblUpperFee)
   End If

curAgentFee_Exit:
   Exit Function
curAgentFee_Error:
   Beep
   MsgBox Err.Description
   Resume curAgentFee_Exit

End Function



'========================= basCalandar
Option Compare Database
Option Explicit

Const adhcCalendarForm = "frmCalendar"

Function adhDoCalendar(Optional varPassedDate As Variant) As Variant
    '
    ' This is the public entry point.
    ' If the passed in date is missing (as it will
    ' be if someone just opens the Calendar form
    ' raw), start on the current day.
    ' Otherwise, start with the date that is passed in.
    
    Dim varStartDate As Variant

    ' If they passed a value at all, attempt to
    ' use it as the start date.
    varStartDate = IIf(IsMissing(varPassedDate), _
     Date, varPassedDate)
    ' OK, so they passed a value that wasn't a date.
    ' Just use today's date in that case, too.
    If Not IsDate(varStartDate) Then varStartDate = Date
    DoCmd.OpenForm FormName:=adhcCalendarForm, _
     WindowMode:=acDialog, OpenArgs:=varStartDate

    ' You won't get here until the form is
    ' closed or hidden.
    '
    ' If the form is still loaded, then get the
    ' final chosen date from the form.  If it isn't,
    ' return Null.
    If isOpen(adhcCalendarForm) Then
        adhDoCalendar = Forms(adhcCalendarForm).Value
        DoCmd.Close acForm, adhcCalendarForm
    Else
        adhDoCalendar = Null
    End If
End Function

Private Function isOpen(strName As String, _
 Optional intObjectType As Integer = acForm)
    ' Returns True if strName is open, False otherwise.
    ' Assume the caller wants to know about a form.
    isOpen = (SysCmd(acSysCmdGetObjectState, _
     intObjectType, strName) <> 0)
End Function

'======================================== basInterpolate
Option Compare Database

'Purpose:      This routine calculates the straight line interpolation of an output range
'              based on the input value's position in the input range.
'
'Parameters:
'  dblRangeValue  in - The value we want an interpolated output value for
'  dblRangeLower  in - The value below the input value
'  dblRangeUpper  in - The value above the input value
'  dblOutputLower in - The output value for lower range value
'  dblOutputUpper in - The output value for upper range value
'  Returns           - An interpolated value between the output values based on the input values
'
'Notes:        This routine assumes that the input value lies between the lower and upper
'              input values. No checking is done. However, the routine will work if the
'              input value lies outside of the input range or even if the lower and upper
'              values are reversed.
'
'History:
'  08/10/97 Zeke Leventhal-Arnold
'     Initial Documentation
'
Function Interpolate(ByVal dblRangeValue As Double, _
         ByVal dblRangeLower As Double, ByVal dblRangeUpper As Double, _
         ByVal dblOutputLower As Double, ByVal dblOutputUpper As Double) As Double

   If dblRangeLower = dblRangeUpper Then
      Interpolate = dblOutputLower     ' Choose either output value even if different
   Else
      Interpolate = ((dblRangeValue - dblRangeLower) / (dblRangeUpper - dblRangeLower) * _
            (dblOutputUpper - dblOutputLower)) + dblOutputLower
   End If

End Function




'================================================ NextDateFwd
Option Compare Database
Option Explicit

Function NextDateFwd(DeliveryYear, DeliveryMonth, WASettleDay, ScheduleType)

If ScheduleType = "MS" Then
   
   If WASettleDay <= 15 And DeliveryMonth <> 12 Then
       NextDateFwd = DateSerial(DeliveryYear, (DeliveryMonth + 1), 18)
   ElseIf WASettleDay <= 15 And DeliveryMonth = 12 Then
       NextDateFwd = DateSerial((DeliveryYear + 1), 1, 18)
   ElseIf WASettleDay > 15 And DeliveryMonth <= 11 Then
       NextDateFwd = DateSerial(DeliveryYear, (DeliveryMonth + 1), 18)
   ElseIf WASettleDay > 15 And DeliveryMonth = 12 Then
       NextDateFwd = DateSerial((DeliveryYear + 1), 1, 18)
   End If

ElseIf ScheduleType = "GL" Then
   
   If WASettleDay <= 15 And DeliveryMonth <> 12 Then
       NextDateFwd = DateSerial(DeliveryYear, (DeliveryMonth + 1), 18)
   ElseIf WASettleDay <= 15 And DeliveryMonth = 12 Then
       NextDateFwd = DateSerial((DeliveryYear + 1), 1, 18)
   ElseIf WASettleDay > 15 And DeliveryMonth <= 11 Then
       NextDateFwd = DateSerial(DeliveryYear, (DeliveryMonth + 1), 18)
   ElseIf WASettleDay > 15 And DeliveryMonth = 12 Then
       NextDateFwd = DateSerial((DeliveryYear + 1), 1, 18)
   End If

ElseIf ScheduleType = "SS" Then
   
   If WASettleDay <= 15 And DeliveryMonth <> 12 Then
       NextDateFwd = DateSerial(DeliveryYear, (DeliveryMonth + 1), 18)
   ElseIf WASettleDay <= 15 And DeliveryMonth = 12 Then
       NextDateFwd = DateSerial((DeliveryYear + 1), 1, 18)
   ElseIf WASettleDay > 15 And DeliveryMonth <= 11 Then
       NextDateFwd = DateSerial(DeliveryYear, (DeliveryMonth + 1), 18)
   ElseIf WASettleDay > 15 And DeliveryMonth = 12 Then
       NextDateFwd = DateSerial((DeliveryYear + 1), 1, 18)
   End If

ElseIf ScheduleType = "MA" Then
   
   If WASettleDay <= 15 And DeliveryMonth <> 12 Then
       NextDateFwd = DateSerial(DeliveryYear, (DeliveryMonth + 1), 2)
   ElseIf WASettleDay <= 15 And DeliveryMonth = 12 Then
       NextDateFwd = DateSerial((DeliveryYear + 1), 1, 2)
   ElseIf WASettleDay > 15 And DeliveryMonth <= 11 Then
       NextDateFwd = DateSerial(DeliveryYear, (DeliveryMonth + 1), 2)
   ElseIf WASettleDay > 15 And DeliveryMonth = 12 Then
       NextDateFwd = DateSerial((DeliveryYear + 1), 1, 2)
   End If

ElseIf ScheduleType = "AA" Then
   
   If WASettleDay <= 15 And DeliveryMonth <= 10 Then
       NextDateFwd = DateSerial(DeliveryYear, (DeliveryMonth + 2), 18)
   ElseIf WASettleDay <= 15 And DeliveryMonth = 11 Then
       NextDateFwd = DateSerial((DeliveryYear + 1), 1, 18)
   ElseIf WASettleDay <= 15 And DeliveryMonth = 12 Then
       NextDateFwd = DateSerial((DeliveryYear + 1), 2, 18)
       
   ElseIf WASettleDay > 15 And DeliveryMonth <= 10 Then
       NextDateFwd = DateSerial(DeliveryYear, (DeliveryMonth + 2), 18)
   ElseIf WASettleDay > 15 And DeliveryMonth = 11 Then
       NextDateFwd = DateSerial((DeliveryYear + 1), 1, 18)
   ElseIf WASettleDay > 15 And DeliveryMonth = 12 Then
       NextDateFwd = DateSerial((DeliveryYear + 1), 2, 18)
   End If

End If
   
   
End Function


'======================================= NextDate
Function NextDate(EntryDate, ScheduleType)

If ScheduleType = "MS" Then
   
   If DatePart("d", EntryDate) >= 18 And DatePart("m", EntryDate) <> 12 Then
       NextDate = DateSerial(Year(EntryDate), Month(EntryDate) + 1, 18)
   ElseIf DatePart("d", EntryDate) >= 18 And DatePart("m", EntryDate) = 12 Then
       NextDate = DateSerial(Year(EntryDate) + 1, 1, 18)
   ElseIf DatePart("d", EntryDate) < 18 And DatePart("m", EntryDate) <> 12 Then
       NextDate = DateSerial(Year(EntryDate), Month(EntryDate) + 1, 18)
   ElseIf DatePart("d", EntryDate) < 18 And DatePart("m", EntryDate) = 12 Then
       NextDate = DateSerial(Year(EntryDate) + 1, 1, 18)
   End If

ElseIf ScheduleType = "GL" Then
   
   If DatePart("d", EntryDate) >= 18 And DatePart("m", EntryDate) <> 12 Then
       NextDate = DateSerial(Year(EntryDate), Month(EntryDate) + 1, 18)
   ElseIf DatePart("d", EntryDate) >= 18 And DatePart("m", EntryDate) = 12 Then
       NextDate = DateSerial(Year(EntryDate) + 1, 1, 18)
   ElseIf DatePart("d", EntryDate) < 18 And DatePart("m", EntryDate) <> 12 Then
       NextDate = DateSerial(Year(EntryDate), Month(EntryDate) + 1, 18)
   ElseIf DatePart("d", EntryDate) < 18 And DatePart("m", EntryDate) = 12 Then
       NextDate = DateSerial(Year(EntryDate) + 1, 1, 18)
   End If

ElseIf ScheduleType = "SS" Then
   
   If DatePart("d", EntryDate) >= 18 And DatePart("m", EntryDate) <> 12 Then
       NextDate = DateSerial(Year(EntryDate), Month(EntryDate) + 1, 18)
   ElseIf DatePart("d", EntryDate) >= 18 And DatePart("m", EntryDate) = 12 Then
       NextDate = DateSerial(Year(EntryDate) + 1, 1, 18)
   ElseIf DatePart("d", EntryDate) < 18 And DatePart("m", EntryDate) <> 12 Then
       NextDate = DateSerial(Year(EntryDate), Month(EntryDate) + 1, 18)
   ElseIf DatePart("d", EntryDate) < 18 And DatePart("m", EntryDate) = 12 Then
       NextDate = DateSerial(Year(EntryDate) + 1, 1, 18)
   End If

ElseIf ScheduleType = "MA" Then
   
   If DatePart("d", EntryDate) >= 2 And DatePart("m", EntryDate) <> 12 Then
       NextDate = DateSerial(Year(EntryDate), Month(EntryDate) + 1, 2)
   ElseIf DatePart("d", EntryDate) >= 2 And DatePart("m", EntryDate) = 12 Then
       NextDate = DateSerial(Year(EntryDate) + 1, 1, 2)
   ElseIf DatePart("d", EntryDate) < 2 Then
       NextDate = DateSerial(Year(EntryDate), Month(EntryDate), 2)
   End If

ElseIf ScheduleType = "AA" Then
   
   If DatePart("d", EntryDate) >= 18 And (DatePart("m", EntryDate) + 1) < 12 Then
       NextDate = DateSerial(Year(EntryDate), Month(EntryDate) + 2, 18)
   ElseIf DatePart("d", EntryDate) >= 18 And (DatePart("m", EntryDate) + 1) = 12 Then
       NextDate = DateSerial(Year(EntryDate) + 1, 1, 18)
   ElseIf DatePart("d", EntryDate) >= 18 And (DatePart("m", EntryDate) + 1) > 12 Then
       NextDate = DateSerial(Year(EntryDate) + 1, 2, 18)
   
   ElseIf DatePart("d", EntryDate) < 18 And (DatePart("m", EntryDate) + 1) < 12 Then
       NextDate = DateSerial(Year(EntryDate), Month(EntryDate) + 2, 18)
   ElseIf DatePart("d", EntryDate) < 18 And (DatePart("m", EntryDate) + 1) = 12 Then
       NextDate = DateSerial(Year(EntryDate) + 1, 1, 18)
   ElseIf DatePart("d", EntryDate) < 18 And (DatePart("m", EntryDate) + 1) > 12 Then
       NextDate = DateSerial(Year(EntryDate) + 1, 2, 18)
   End If

End If
   
   
End Function

'================================== mod_enableShiftKey
Public Sub EnableShiftKeyByPass()

On Error GoTo err_handler:

    Set prp = CurrentDb.CreateProperty("AllowBypassKey", _
    dbBoolean, True, True)
    CurrentDb.Properties.Append prp
    Set prp = Nothing
    
    MsgBox "Database is no longer locked"

Exit Sub

err_handler:

If Err.Number = 3367 Then

    CurrentDb.Properties("AllowBypassKey") = True
    Resume Next

Else

    MsgBox "Error:  " & Err.Number & vbCrLf & Err.Description, vbCritical

End If

End Sub

'============================== getQueryDef
Option Compare Database
Option Explicit

Dim QueryName As String


Public Sub F1()
    On Error GoTo eh
    Dim db As Database
    Dim q As QueryDef
    QueryName = "qryAppendMPFForwardCommitmentsToInstrumentP"
    Set db = DBEngine.Workspaces(0).Databases(0)
    
    'QueryName = "qryAppendMPFSecuritiesToInstrumentTableP"
    'QueryName = "qryAppendMPFSecuritiesToPortfolioContentsTableP"
    'QueryName = "qryAppendMPFForwardCommitmentsToInstrumentP"
    QueryName = "qryAppendMPFForwardCommitmentsToPortfolioContentsP"
    
    For Each q In db.QueryDefs
        If q.Name = QueryName Then
            Debug.Print QueryName
            Debug.Print q.SQL
        End If
    Next
    
    Exit Sub
eh:
    MsgBox Err.Description, vbInformation
End Sub

'================================== shiftKeyByPass
Option Compare Database   'Use database order for string comparisons
Public Sub DisableShiftKeyBypass()

On Error GoTo err_handler:

    Set prp = CurrentDb.CreateProperty("AllowBypassKey", _
    dbBoolean, False, True)
    CurrentDb.Properties.Append prp
    Set prp = Nothing
    
    MsgBox "Database is locked"

Exit Sub

err_handler:

If Err.Number = 3367 Then

    CurrentDb.Properties("AllowBypassKey") = False
    Resume Next

Else

    MsgBox "Error:  " & Err.Number & vbCrLf & Err.Description, vbCritical

End If

End Sub


'================= modAudit
Option Compare Database
Option Explicit

Public Function InsertFromAudit(frmForm As Form, sTableName As String)

On Error GoTo err_handler

Dim objDAO          As DAO.Recordset
Dim fControl        As Object
Dim sSQL            As String
Dim sSQLColumns     As String
Dim sSQLValues      As String

    For Each fControl In frmForm.Controls

        If fControl.ControlType = 109 Then
        
            If fControl.Name <> "UserID" And fControl.Name <> "ActionTaken" And fControl.Name <> "ActionTime" Then
            
                sSQLColumns = sSQLColumns & "[" & fControl.Name & "]" & ", "
                
                If IsNull(fControl.Value) Then
                    sSQLValues = sSQLValues & "NULL, "
                ElseIf IsNumeric(fControl.Value) And Left(fControl.Value, 1) <> "0" Then
                    sSQLValues = sSQLValues & fControl.Value & ", "
                Else
                    sSQLValues = sSQLValues & "'" & fControl.Value & "', "
                End If
                
            End If
        
                        
        End If
        
    Next
            
sSQL = "Insert into " & sTableName & "(" & Left(sSQLColumns, Len(sSQLColumns) - 2) & ") VALUES ( " & Left(sSQLValues, Len(sSQLValues) - 2) & ")"
MsgBox sSQL
CurrentDb.Execute sSQL

InsertFromAudit = 0

Exit Function

err_handler:

MsgBox Err.Description, vbCritical, "Error Encountered"
InsertFromAudit = Err.Number

End Function

'================================ RefreshLinks
Option Compare Database
Option Explicit

Public Function RefreshLinks()

Dim tdf As TableDef

On Error GoTo err_handler

For Each tdf In CurrentDb.TableDefs

    If InStr(tdf.Connect, "DSN=MPFDW") Then
    
        tdf.Connect = "DSN=MPFFHLBDW"
        tdf.RefreshLink
    
    End If

Next


Exit Function

err_handler:
MsgBox Err.Description

End Function

'===================== modSecurity
Option Compare Database
Option Explicit

     ' Declare for call to mpr.dll.
   Declare Function WNetGetUser Lib "mpr.dll" _
      Alias "WNetGetUserA" (ByVal lpName As String, _
      ByVal lpUserName As String, lpnLength As Long) As Long

   Const NoError = 0       'The Function call was successful

Public Function GetUserName() As String
    
    ' Buffer size for the return string.
    Const lpnLength As Integer = 255
    
    ' Get return buffer space.
    Dim status As Integer
    
    ' For getting user information.
    Dim lpName, lpUserName As String
    
    ' Assign the buffer size constant to lpUserName.
    lpUserName = Space$(lpnLength + 1)
    
    ' Get the log-on name of the person using product.
    status = WNetGetUser(lpName, lpUserName, lpnLength)
    
    ' See whether error occurred.
    If status = NoError Then
       ' This line removes the null character. Strings in C are null-
       ' terminated. Strings in Visual Basic are not null-terminated.
       ' The null character must be removed from the C strings to be used
       ' cleanly in Visual Basic.
       lpUserName = Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1)
    Else
    
       ' An error occurred.
       MsgBox "Unable to get the name."
       End
    End If
    
    ' Display the name of the person logged on to the machine.
    GetUserName = lpUserName

End Function

Public Function CheckAdminRights() As Boolean

    
    Dim rsMYTable   As DAO.Recordset
    Dim sSQL        As String
    Dim RetVal      As Boolean

    'Find the Sum of the Trades being hedged for calculating the percentage of the hedge to apply.
    sSQL = "Select count(*) as rCount from FHLBNAV_Admin where UserName = '" & GetUserName() & "'"
            
    Set rsMYTable = CurrentDb.OpenRecordset(sSQL)
    
    If rsMYTable("rCount") = 0 Then

        RetVal = False

    Else
    
        RetVal = True
    
    End If

    rsMYTable.Close

CheckAdminRights = RetVal

End Function




