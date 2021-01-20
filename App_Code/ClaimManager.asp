<%
Class CClaimManager
	Private mClaim
	Private mList
	Private mMessages
	Private mConnection
	Private mErrorNumber
	
	'Constructor
	Private Sub Class_Initialize()
		On Error Resume Next

		mMessages=""
		mErrorNumber=0
		Set mConnection = Server.CreateObject("ADODB.Connection")
		mConnection.Open ConnectionString("MedicalOld")
		' mConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("db\Reimbursements.mdb") & ";User Id=admin;Password=;"
		'mConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("db\WordDefs.accdb") & ";Persist Security Info=False;"
		CheckForError "Open Connection Failed!"
		AddMessage "Connection State: " & mConnection.State, true
		Set mClaim=Nothing
		Set mList=Nothing
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(mConnection) Then
			If mConnection.State = 1 Then
				mConnection.Close
			End If
			Set mConnection = Nothing
		End If
		
		If IsObject(mList) Then Set mList = Nothing
	End Sub

    Private Sub CheckForError(message)
        If Err.number > 0 Then
			mErrorNumber=Err.number
            SetCustomError message
            Err.Clear
        End If
    End Sub

    ' Sub SetCustomError(strMessage)
        ' Display custom message and information from VBScript Err object.

        ' AddMessage "<br/>" & strMessage & "<br/>" & _
          ' "Number (dec) : " & Err.Number & "<br/>" & _
          ' "Number (hex) : &H" & Hex(Err.Number) & "<br/>" & _
          ' "Description  : " & Err.Description & "<br/>" & _
          ' "Source       : " & Err.Source, true
        ' Err.Clear
    ' End Sub

 	'Use this for debugging AddMessage "message", true
	' Private Sub AddMessage(message, add)
		' If add Then
			' mMessages=mMessages & message & "<br/>"
		' End If
	' End Sub

    ' Function StripApostrophe(text)
        ' StripApostrophe = replace (text,"'","&acute;")
    ' End Function

	' Private Function AddArgument(value, argType)
		' Dim myValue
		
	    ' Select Case argType
	        ' Case "Numeric"
	            ' myValue = value & ", "
	        ' Case "NumericTerminal"
	            'Nothing to do just return the number
				' myValue = value
	        ' Case "Text"
	            ' myValue = "'" & StripApostrophe(value) & "', "
	        ' Case "TextTerminal"
	            ' myValue = "'" & StripApostrophe(value) & "'"
			' Case "Date"
				' myValue = "#" & value & "#, "
			' Case "DateTerminal"
				' myValue = "#" & value & "#"
	    ' End Select
	    
	    ' AddArgument=myValue
	' End Function
	
	' Private Function GetFieldValue(myRS, field, fieldType)
	    ' Dim value
	    
	    ' Select Case fieldType
	        ' Case "Numeric"
	            ' If IsNull(myRS(field)) Then value=0 Else value=myRS(field) End If
	        ' Case Else '"Text"
	            ' If IsNull(myRS(field)) Then value="" Else value=Trim(myRS(field)) End If
	    ' End Select    
	    
	    ' GetFieldValue=value
	' End Function	

	Private Function FillClaim(rs)
		On Error Resume Next
		Dim value
		
		Set value=New CClaim

		value.ClaimID=GetFieldValue(rs,"ClaimID","Numeric")
		value.PatientID=GetFieldValue(rs,"PatientID","Numeric")
		value.PatientFirstName=GetFieldValue(rs,"PatientFirstName","Text")
		value.PatientMi=GetFieldValue(rs,"PatientMi","Text")
		value.PatientLastName=GetFieldValue(rs,"PatientLastName","Text")
		value.ProviderID=GetFieldValue(rs,"ProviderID","Numeric")
		value.ProviderFirstName=GetFieldValue(rs,"ProviderFirstName","Text")
		value.ProviderMi=GetFieldValue(rs,"ProviderMi","Text")
		value.ProviderLastName=GetFieldValue(rs,"ProviderLastName","Text")
		value.ServiceID=GetFieldValue(rs,"ServiceID","Numeric")
		value.ServiceName=GetFieldValue(rs,"ServiceName","Text")
		value.ExpenseTypeID=GetFieldValue(rs,"ExpenseTypeID","Numeric")
		value.ExpenseTypeName=GetFieldValue(rs,"ExpenseTypeName","Text")
		value.ExpenseDate=GetFieldValue(rs,"ExpenseDate","Text")
		value.ExpenseAmount=GetFieldValue(rs,"ExpenseAmount","Numeric")
		value.ClaimNumber=GetFieldValue(rs,"ClaimNumber","Text")
		value.InsuranceClaimNumber=GetFieldValue(rs,"InsuranceClaimNumber","Text")
		value.PaidCD=GetFieldValue(rs,"PaidCD","Numeric")
		value.MedicationID=GetFieldValue(rs,"MedicationID","Numeric")
		value.MedicationName=GetFieldValue(rs,"MedicationName","Text")
		value.MedicationAmount=GetFieldValue(rs,"MedicationAmount","Numeric")
		value.PayeeID=GetFieldValue(rs,"PayeeID","Numeric")
		value.PayeeName=GetFieldValue(rs,"PayeeName","Text")
		value.InvoiceID=GetFieldValue(rs,"InvoiceID","Numeric")
		value.InvoiceNumber=GetFieldValue(rs,"InvoiceNumber","Text")
		
		Set FillClaim = value	
	End Function
	
	Private Function FillPaidClaimAmount(rs)
		Dim value
		
		Set value=New CClaim
		value.InsuranceClaimNumber=GetFieldValue(rs,"InsuranceClaimNumber","Numeric")
		value.ExpenseAmount=GetFieldValue(rs,"ExpenseAmount","Numeric")
	
		Set FillPaidClaimAmount = value
	End Function

	Private Function FillMonthlyClaim(rs)
		Dim value
		
		Set value=New CMonthlyClaim
		value.MonthID=GetFieldValue(rs,"MonthID","Numeric")
		value.Total=GetFieldValue(rs,"Total","Numeric")
	
		Set FillMonthlyClaim = value
	End Function
	
	Private Function FillClaimNumber(rs)
		Dim value
		
		Set value=New CClaim

		value.ClaimNumber=GetFieldValue(rs,"ClaimNumber","Text")
		
		Set FillClaimNumber = value
	End Function

	Private Function FillInsuranceClaimNumber(rs)
		Dim value
		
		Set value=New CClaim

		value.InsuranceClaimNumber=GetFieldValue(rs,"InsuranceClaimNumber","Text")
		
		Set FillInsuranceClaimNumber = value
	End Function
	
	Private Function FillClaimAmount(rs)
		Dim value
		
		Set value=New CClaim

		value.ServiceName=GetFieldValue(rs,"ServiceName","Text")
		value.ExpenseAmount=GetFieldValue(rs,"ExpenseAmount","Numeric")
		
		Set FillClaimAmount = value
	End Function
	
	Private Function FillFlexPayRequestAmount(rs)
		Dim value
		
		Set value=New CFlexPayRequestAmount

		value.ExpenseType=GetFieldValue(rs,"ExpenseTypeName","Text")
		value.StartDate=GetFieldValue(rs,"StartDate","Text")
		value.EndDate=GetFieldValue(rs,"EndDate","Text")
		value.ExpenseAmount=GetFieldValue(rs,"Expense","Numeric")
		
		Set FillFlexPayRequestAmount = value
	End Function
	'Public Properties
	Public Property Set Claim(value)
		Set mClaim=value
	End Property
	
	Public Property Get List()
		Set List=mList
	End Property
	
	Public Property Get Messages()
		Messages=mMessages
	End Property
	
	'Public Methods
	Public Sub SelectClaimNumbers()
	    On Error Resume Next
		Dim value
		Dim rs

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimNumbers", mConnection
		CheckForError "Select Claim Numbers Failed!"

		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaimNumber(rs)
			mList.Add CStr(value.ClaimNumber), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
		Set value = Nothing
	End Sub

	Public Sub SelectInsuranceClaimNumbers()
	    On Error Resume Next
		Dim value
		Dim rs

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInsuranceClaimNumbers", mConnection
		CheckForError "Select Insurance Claim Numbers Failed!"

		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInsuranceClaimNumber(rs)
			mList.Add CStr(value.InsuranceClaimNumber), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
		Set value = Nothing
	End Sub

	Public Sub SelectClaims()
	    On Error Resume Next
		Dim value
		Dim rs

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaims", mConnection
		CheckForError "Select Claims Failed!"

		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaim(rs)
			mList.Add CStr(value.ClaimID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
		Set value = Nothing
	End Sub

	Public Sub SelectClaimsByAmount(amount)
	    On Error Resume Next
		Dim value
		Dim rs

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimsByAmount " & _
		AddArgument(amount, "NumericTerminal"), mConnection
		CheckForError "Select Claims Failed!"

		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaim(rs)
			mList.Add CStr(value.ClaimID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
		Set value = Nothing
	End Sub

	Public Sub SelectPaidClaimAmountsByDate(startDate, endDate)
	    On Error Resume Next
		Dim value
		Dim rs

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelPaidInsuranceClaimAmountsByDate " & _
			AddArgument(startDate, "Date") & _
			AddArgument(endDate, "DateTerminal"), mConnection
		CheckForError "Select Paid Claim Amounts by dates Failed!"

		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillPaidClaimAmount(rs)
			mList.Add CStr(value.InsuranceClaimNumber), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
		Set value = Nothing
	End Sub

	Public Sub SelectClaimsByDates(startDate, endDate)
	    'On Error Resume Next
		Dim value
		Dim rs

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimsByDates " & _
			AddArgument(startDate, "Date") & _
			AddArgument(endDate, "DateTerminal"), mConnection
		CheckForError "Select Claims by dates Failed!"

		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaim(rs)
			mList.Add CStr(value.ClaimID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
		Set value = Nothing
	End Sub
	
	Public Sub SelectClaimsByMonth(startDate, endDate, claimType)
		On Error Resume Next
		Dim value
		Dim rs
		
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimsByMonth " & _
		AddArgument(startDate, "Date") & _
		AddArgument(endDate, "Date") & _
		AddArgument(claimType, "NumericTerminal"), mConnection
		
		CheckForError "Select Claims by Month Failed!"
		
		Set mList = CreateObject("Scripting.Dictionary")
		If rs.State = 1 Then
			While Not rs.EOF
				Set value = FillMonthlyClaim(rs)
				mList.Add CStr(value.MonthID), value
				rs.MoveNext
			Wend
		End If
		
		rs.Close
		Set rs = Nothing
		Set value = Nothing
	End Sub
	
	Public Sub SelectUnpaidClaimsByDates(startDate, endDate)
	    'On Error Resume Next
		Dim value
		Dim rs

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelUnpaidClaimsByDates " & _
			AddArgument(startDate, "Date") & _
			AddArgument(endDate, "DateTerminal"), mConnection
		CheckForError "Select Unpaid Claims by dates Failed!"

		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaim(rs)
			mList.Add CStr(value.ClaimID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing
		Set value = Nothing
	End Sub

	Public Sub SelectClaimByID(ClaimID)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimByID " & _
			AddArgument(ClaimID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Claim by ID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaim(rs)
			mList.Add value.ClaimIDToString, value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
		
	Public Sub SelectClaimsByPatientID(patientID)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimsByPatientID " & _
			AddArgument(patientID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Claim by patientID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaim(rs)
			mList.Add CStr(value.ClaimID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub

	Public Sub SelectClaimsByPayeeID(payeeID)
	    'On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimsByPayeeID " & _
			AddArgument(payeeID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Claim by payeeID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaim(rs)
			mList.Add CStr(value.ClaimID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub

	Public Sub SelectClaimsWithoutInvoiceByPayeeID(payeeID)
	    'On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimsWithoutInvoiceByPayeeID " & _
			AddArgument(payeeID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Claim by payeeID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaim(rs)
			mList.Add CStr(value.ClaimID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub

	Public Sub SelectClaimsByProviderID(providerID)
	    'On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimsByProviderID " & _
			AddArgument(providerID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Claim by providerID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaim(rs)
			mList.Add CStr(value.ClaimID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub

	Public Sub SelectClaimsByServiceID(serviceID)
	    'On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimsByServiceID " & _
			AddArgument(serviceID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Claim by serviceID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaim(rs)
			mList.Add CStr(value.ClaimID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub

	Public Sub SelectClaimsByServiceIDAndDate(serviceID, startDate, endDate)
	    'On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimsByServiceIDAndDate " & _
			AddArgument(serviceID, "Numeric") & _
			AddArgument(startDate, "Date") & _
			AddArgument(endDate, "DateTerminal"), mConnection
		CheckForError "Unable to select the Claim by serviceID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaim(rs)
			mList.Add CStr(value.ClaimID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub

	Public Sub SelectClaimsByClaimNumber(claimNumber)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimsByClaimNumber " & _
			AddArgument(claimNumber, "TextTerminal"), mConnection
		CheckForError "Unable to select the Claim by claim number!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaim(rs)
			mList.Add CStr(value.ClaimID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub

	Public Sub SelectClaimsByInsuranceClaimNumber(claimNumber)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimsByInsuranceClaimNumber " & _
			AddArgument(claimNumber, "TextTerminal"), mConnection
		CheckForError "Unable to select the Claim by insurnace claim number!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaim(rs)
			mList.Add CStr(value.ClaimID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub

	Public Sub SelectClaimAmountByDates(startDate, endDate)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelClaimAmountByDates " & _
			AddArgument(startDate, "Date") & _
			AddArgument(endDate, "DateTerminal"), mConnection
		CheckForError "Unable to select the Claim amount by dates!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillClaimAmount(rs)
			mList.Add CStr(value.ServiceName), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
	
	Public Sub SelectFlexPayRequestAmountByDates(startDate, endDate)
	    'On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelFlexPayRequestAmounts " & _
			AddArgument(startDate, "Date") & _
			AddArgument(endDate, "DateTerminal"), mConnection
		CheckForError "Unable to select the Flex Pay amounts by dates!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillFlexPayRequestAmount(rs)
			mList.Add value.ExpenseType, value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
	
	Public Sub SelectFlexPayRequestAmountsByInsuranceClaimNumber(claimNumber)
	    'On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelFlexPayRequestAmountsByInsuranceClaimNumber " & _
			AddArgument(claimNumber, "TextTerminal"), mConnection
		CheckForError "Unable to select the Flex Pay amounts by insurance claim number!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillFlexPayRequestAmount(rs)
			value.InsuranceClaimNumber = claimNumber
			mList.Add value.ExpenseType, value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub

	Public Function Delete()
		On Error Resume Next
		Dim sql
		
		sql = "pDelClaim " & _
		AddArgument(mClaim.ClaimID, "NumericTerminal")
		
		mConnection.Execute sql
		CheckForError "Unable to delete Claim!"
		If mErrorNumber = 0 Then
		    Delete=True
		Else
		    Delete=False
		End If
	End Function
	
	'Microsoft JET Database Engine (0x80004005) Operation must use an updateable query.
	'Modified the permissions for all users to have full control
	Public Function Save()
		On Error Resume Next
		Dim sql

		sql = "pInsClaim " & _
		AddArgument(mClaim.PatientID, "Numeric") & _
		AddArgument(mClaim.ProviderID, "Numeric") & _
		AddArgument(mClaim.ServiceID, "Numeric") & _
		AddArgument(mClaim.ExpenseDate, "Text") & _
		AddArgument(mClaim.ExpenseAmount, "Numeric") & _
		AddArgument(mClaim.ExpenseTypeID, "Numeric") & _
		AddArgument(mClaim.ClaimNumber, "Text") & _
		AddArgument(mClaim.InsuranceClaimNumber, "Text") & _
		AddArgument(mClaim.PaidCD, "Numeric") & _
		AddArgument(mClaim.MedicationID, "Numeric") & _
		AddArgument(mClaim.MedicationAmount, "NumericTerminal")
		mConnection.Execute sql
		CheckForError "Unable to save new Claim!"
		If mErrorNumber = 0 Then
		    Save=True
		Else
		    Save=False
		End If		
	End Function

	Public Function Update()
		' On Error Resume Next
		Dim sql
		
		sql = "pUpdClaim " & _
		AddArgument(mClaim.ClaimID, "Numeric") & _
		AddArgument(mClaim.PatientID, "Numeric") & _
		AddArgument(mClaim.ProviderID, "Numeric") & _
		AddArgument(mClaim.ServiceID, "Numeric") & _
		AddArgument(mClaim.ExpenseDate, "Text") & _
		AddArgument(mClaim.ExpenseAmount, "Numeric") & _
		AddArgument(mClaim.ExpenseTypeID, "Numeric") & _
		AddArgument(mClaim.ClaimNumber, "Text") & _
		AddArgument(mClaim.InsuranceClaimNumber, "Text") & _
		AddArgument(mClaim.PaidCD, "Numeric") & _
		AddArgument(mClaim.MedicationID, "Numeric") & _
		AddArgument(mClaim.MedicationAmount, "Numeric") & _
		AddArgument(mClaim.PayeeID, "NumericTerminal")
		
		mConnection.Execute sql
		CheckForError "Unable to update Claim!"
		If mErrorNumber = 0 Then
		    Update=True
		Else
		    Update=False
		End If
	End Function
End Class
%>