<!-- #include file="PayeeManager.asp" -->
<!-- #include file="Payee.asp" -->
<!-- #include file="PatientManager.asp" -->
<!-- #include file="Patient.asp" -->
<!-- #include file="ProviderManager.asp" -->
<!-- #include file="Provider.asp" -->
<!-- #include file="ServiceManager.asp" -->
<!-- #include file="Service.asp" -->
<!-- #include file="ClaimManager.asp" -->
<!-- #include file="Claim.asp" -->
<!-- #include file="MonthlyClaim.asp" -->
<!-- #include file="MedicationManager.asp" -->
<!-- #include file="Medication.asp" -->
<!-- #include file="ExpenseTypeManager.asp" -->
<!-- #include file="ExpenseType.asp" -->
<!-- #include file="FlexPayRequestAmount.asp" -->
<!-- #include file="Common.asp" -->
<%
Class CClaimPageManager
	Private mMessages

 	'Use this for debugging AddMessage "message", true
	Private Sub AddMessage(message, add)
		If add Then
			mMessages=mMessages & message & "<br/>"
		End If
	End Sub
	
	Private Function LoadList(list)
		Dim i
		Dim cls
		Dim keys
		Dim value
		Dim output
		Dim total
		Dim rowCount
		
		keys=list.Keys
		total=0
		rowCount=0
		
		output="<table>"
		output=output & "<tr>"
		output=output & "<th>Date</th>"
		output=output & "<th>Patient</th>"
		output=output & "<th>Service</th>"
		output=output & "<th>Provider</th>"
		output=output & "<th>Amount</th>"
		output=output & "<th>Claim ID</th>"
		output=output & "<th>Insurance ID</th>"
		output=output & "<th>Invoice ID</th>"
		output=output & "<th>Paid</th>"
		output=output & "<th>Medication</th>"
		output=output & "<th>Medication<br/>Amount</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			Set value=list.Item(keys(i))
			total = total + CDbl(value.ExpenseAmount)
			
			output=output & "<tr>"
			output=output & "<td>" & value.ExpenseDate & "</td>"
			output=output & "<td>" & value.PatientName & "</td>"
			output=output & "<td>" & value.ServiceName & "</td>"
			output=output & "<td>" & value.ProviderName & "</td>"
			output=output & "<td align='right'>" & FormatNumber(value.ExpenseAmount, 2) & "</td>"
			output=output & "<td align='right'><a href=javascript:popUpBig('Claims/" & value.ClaimNumber & ".pdf');>" & value.ClaimNumber & "</a></td>"
			output=output & "<td align='right'>" & value.InsuranceClaimNumber & "</td>"
			output=output & "<td align='right'>"
			output=output & IIf(value.PayeeID = 0, "&nbsp;", "<a href=javascript:popUpBig('Invoices/" & value.InvoiceNumber & ".pdf');>" & value.InvoiceNumber & "</a>")
			output=output & "</td>"
			output=output & "<td>" & value.Paid & "</td>"
			output=output & "<td>" & value.MedicationName & "</td>"
			output=output & "<td align='center'>" & IIf(CDbl(value.MedicationAmount) = 0, "&nbsp;",IIf(CDbl(value.MedicationAmount) < 1, FormatNumber(value.MedicationAmount, 1), CInt(value.MedicationAmount))) & "</td>"
			'output=output & "<td align='center'>" & IIf(CDbl(value.MedicationAmount) = 0, "&nbsp;",FormatNumber(0.0, 2)) & "</td>"
			output=output & "<td>" 
			output=output & "<a href='EditClaim.asp?ClaimID=" & value.ClaimID & "' class='image' title=''><img alt='Edit' src='images/pencil.svg' width='16' height='16' border='0' /></a>"
			output=output & "</td>"
			output=output & "<td>" 
			output=output & "<a onclick='return confirmSubmit()' href='DeleteClaimPr.asp?ClaimID=" & value.ClaimID & "' class='image' title=''><img alt='Edit' src='images/x.svg' width='16' height='16' border='0' /></a>"
			output=output & "</td>"
			output=output & "</tr>"
			
			rowCount = rowCount + 1
		Next

		output=output & "<tr>"
		output=output & "<th>" & rowCount & "</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th class='align_right'>" & FormatNumber(total, 2) & "</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "</tr>"

		Set value=Nothing
		
		output=output & "</table>"
		
		LoadList = output
	End Function
	
	Private Function LoadClaimAmounts(list)
		Dim i
		Dim cls
		Dim keys
		Dim value
		Dim output
		Dim total
		
		keys=list.Keys
		total=0
		rowCount=0
		
		output="<table class='withborder' cellpadding='1' cellspacing='1'>"
		output=output & "<tr>"
		output=output & "<th>Service</th>"
		output=output & "<th>Amount</th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			cls = IIf(cls="rowalt", "row", "rowalt")
			Set value=list.Item(keys(i))
			total = total + CDbl(value.ExpenseAmount)
			
			output=output & "<tr class='" & cls & "'>"
			output=output & "<td>" & value.ServiceName & "</td>"
			output=output & "<td align='right'>" & FormatNumber(value.ExpenseAmount, 2) & "</td>"
			output=output & "</tr>"
			
		Next
		output=output & "<tr>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th align='right'>" & FormatNumber(total, 2) & "</th>"
		output=output & "</tr>"

		Set value=Nothing
		
		output=output & "</table>"
		
		LoadClaimAmounts = output
	End Function
	
	Private Function LoadFlexPayRequestAmounts(list)
		Dim i
		Dim cls
		Dim keys
		Dim value
		Dim output
		Dim total
		
		keys=list.Keys
		total=0
		
		output="<table class='withborder' cellpadding='1' cellspacing='1'>"
		output=output & "<tr>"
		output=output & "<th>Expense Type</th>"
		output=output & "<th>Start<br/>Date</th>"
		output=output & "<th>End<br/>Date</th>"
		output=output & "<th>Expense</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			cls = IIf(cls="rowalt", "row", "rowalt")
			Set value=list.Item(keys(i))
			total = total + CDbl(value.ExpenseAmount)
			
			output=output & "<tr class='" & cls & "'>"
			output=output & "<td>" & value.ExpenseType & "</td>"
			output=output & "<td>" & value.StartDate & "</td>"
			output=output & "<td>" & value.EndDate & "</td>"
			output=output & "<td align='right'>" & FormatNumber(value.ExpenseAmount, 2) & "</td>"
			output=output & "<td><a href=ClaimsIDView.asp?InsuranceClaimNumber=" & value.InsuranceClaimNumber & " class='image' title='' target='ClaimsIDView'><img alt='View' src='images/binoculars.svg' width='16' height='16' border='0' /></a></td>"
			output=output & "</tr>"
			
		Next
		output=output & "<tr>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th class='align_right'>" & FormatNumber(total, 2) & "</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "</tr>"

		Set value=Nothing
		
		output=output & "</table>"
		
		LoadFlexPayRequestAmounts = output
	End Function

	Private Function LoadMonthlyClaims(list)
		Dim i
		Dim cls
		Dim keys
		Dim value
		Dim output
		Dim total
		
		keys=list.Keys
		total=0
		
		output="<table class='withborder' cellpadding='1' cellspacing='1'>"
		output=output & "<tr>"
		output=output & "<th>Month</th>"
		output=output & "<th>Amount</th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			cls = IIf(cls="rowalt", "row", "rowalt")
			Set value=list.Item(keys(i))
			total = total + CDbl(value.Total)
			
			output=output & "<tr class='" & cls & "'>"
			output=output & "<td>" & value.Month & "</td>"
			output=output & "<td class='align_right'>" & FormatNumber(value.Total, 2) & "</td>"
			output=output & "</tr>"
			
		Next
		
		output=output & "<tr>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th class='align_right'>" & FormatNumber(total, 2) & "</th>"
		output=output & "</tr>"

		Set value=Nothing
		
		output=output & "</table>"
		
		LoadMonthlyClaims = output
	End Function

	Private Function LoadPaidClaimAmounts(list)
		Dim i
		Dim cls
		Dim keys
		Dim value
		Dim output
		Dim total
		
		keys=list.Keys
		total=0
		
		output="<table class='withborder' cellpadding='1' cellspacing='1'>"
		output=output & "<tr>"
		output=output & "<th>Insucance<br/>Number</th>"
		output=output & "<th>Amount</th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			cls = IIf(cls="rowalt", "row", "rowalt")
			Set value=list.Item(keys(i))
			total = total + CDbl(value.ExpenseAmount)
			
			output=output & "<tr class='" & cls & "'>"
			output=output & "<td>" & value.InsuranceClaimNumber & "</td>"
			output=output & "<td class='align_right'>" & FormatNumber(value.ExpenseAmount, 2) & "</td>"
			output=output & "</tr>"
			
		Next
		output=output & "<tr>"
		output=output & "<th>&nbsp;</th>"
		output=output & "<th class='align_right'>" & FormatNumber(total, 2) & "</th>"
		output=output & "</tr>"

		Set value=Nothing
		
		output=output & "</table>"
		
		LoadPaidClaimAmounts = output
	End Function
	
	Public Property Get Messages()
		Messages=mMessages
	End Property

	'Public Methods
	Public Function IIf(expression, trueValue, falseValue)
		If expression Then
			IIf = trueValue
		Else
			IIf = falseValue
		End If
	End Function
	
	Public Function LoadClaimTypes(claimType)
		On Error Resume Next
		Dim output

		output="<select name='ClaimType' id='ClaimType'>"	
		output=output & "<option value='1'" & IIf(claimType = "1", " selected", "") & ">Paid</option>"
		output=output & "<option value='0'" & IIf(claimType = "0", " selected", "") & ">Unpaid</option>"
		output=output & "</select>"
				
		LoadClaimTypes=output
	End Function
	
	Public Function LoadClaimNumbers(claimNumber)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CClaimManager		
		manager.SelectClaimNumbers
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
		output="<select name='ClaimNumber'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ClaimNumber = claimNumber, " selected ", "")
		    output=output & "<option value='" & value.ClaimNumber & "'" & selected & ">" & value.ClaimNumber & "</option>"
		Next

		output=output & "</select>"
				
		LoadClaimNumbers=output
		
		Set manager=Nothing
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function LoadClaimNumbersWithOnChange(claimNumber)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CClaimManager		
		manager.SelectClaimNumbers
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, true
		
		keys=list.Keys
		
		output="<select name='ClaimNumber' onChange='javascript:submitform();'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ClaimNumber = claimNumber, " selected ", "")
		    output=output & "<option value='" & value.ClaimNumber & "'" & selected & ">" & value.ClaimNumber & "</option>"
		Next

		output=output & "</select>"
				
		LoadClaimNumbersWithOnChange=output
		
		Set manager=Nothing
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function LoadInsuranceClaimNumbers(claimNumber)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CClaimManager		
		manager.SelectInsuranceClaimNumbers
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
		output="<select name='InsuranceClaimNumber'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.InsuranceClaimNumber = claimNumber, " selected ", "")
		    output=output & "<option value='" & value.InsuranceClaimNumber & "'" & selected & ">" & value.InsuranceClaimNumber & "</option>"
		Next

		output=output & "</select>"
				
		LoadInsuranceClaimNumbers=output
		
		Set manager = Nothing
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function LoadInsuranceClaimNumbersWithOnChange(claimNumber)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CClaimManager		
		manager.SelectInsuranceClaimNumbers
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
		output="<select name='InsuranceClaimNumber' onChange='javascript:submitform();'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.InsuranceClaimNumber = claimNumber, " selected ", "")
		    output=output & "<option value='" & value.InsuranceClaimNumber & "'" & selected & ">" & value.InsuranceClaimNumber & "</option>"
		Next

		output=output & "</select>"
				
		LoadInsuranceClaimNumbersWithOnChange=output
		
		Set manager = Nothing
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function LoadMedications(medicationID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CMedicationManager		
		manager.SelectMedications
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
		output="<select name='MedicationID'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.MedicationID = CInt(medicationID), " selected ", "")
		    output=output & "<option value='" & value.MedicationID & "'" & selected & ">" & value.MedicationName & "</option>"
		Next

		output=output & "</select>"
				
		LoadMedications=output
		
		Set manager = Nothing
		Set list=Nothing
		Set value=Nothing	
	End Function
	
	'
	Public Function LoadMedicationsByProviderID(providerID, medicationID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CMedicationManager		
		manager.SelectMedicationsByProviderID providerID
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
		output="<select name='MedicationID'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.MedicationID = CInt(medicationID), " selected ", "")
		    output=output & "<option value='" & value.MedicationID & "'" & selected & ">" & value.MedicationName & "</option>"
		Next

		output=output & "</select>"
				
		LoadMedicationsByProviderID=output
		
		Set manager = Nothing
		Set list=Nothing
		Set value=Nothing	
	End Function
		
	Public Function LoadPatients(patientID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CPatientManager		
		manager.SelectPatients
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
		output="<select name='PatientID'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.PatientID = CInt(patientID), " selected ", "")
		    output=output & "<option value='" & value.PatientID & "'" & selected & ">" & value.FirstName & " " & IIf(value.Mi <> "", value.Mi & " ", " ") & value.LastName & "</option>"
		Next

		output=output & "</select>"
				
		LoadPatients=output
		
		Set manager=Nothing
		Set list=Nothing
		Set value=Nothing
	End Function
				
	Public Function LoadPatientsWithOnChange(patientID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CPatientManager		
		manager.SelectPatients
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
		output="<select name='PatientID' id='PatientID' onChange='javascript:submitform();'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.PatientID = CInt(patientID), " selected ", "")
		    output=output & "<option value='" & value.PatientID & "'" & selected & ">" & value.FirstName & " " & IIf(value.Mi <> "", value.Mi & " ", " ") & value.LastName & "</option>"
		Next

		output=output & "</select>"
				
		LoadPatientsWithOnChange=output
		
		Set manager=Nothing
		Set list=Nothing
		Set value=Nothing
	End Function
				
	Public Function LoadPayees(payeeID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CPayeeManager		
		manager.SelectPayees
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
	    output="<select name='PayeeID' id='PayeeID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(CInt(value.PayeeID) = CInt(payeeID), " selected ", "")
		    output=output & "<option value='" & value.PayeeID & "'" & selected & ">" & value.PayeeName & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadPayees=output
		
		Set manager = Nothing
		Set value = Nothing
		Set list = Nothing
	End Function
	
	Public Function LoadProviders(providerID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CProviderManager		
		manager.SelectProviders
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
	    output="<select name='ProviderID' id='ProviderID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ProviderID = CInt(providerID), " selected ", "")
		    output=output & "<option value='" & value.ProviderID & "'" & selected & ">" & value.FirstName & " " & IIf(value.Mi <> "", value.Mi & " ", " ") &  value.LastName & "</option>"
	    Next
    	
	    output=output & "</select>"
				
		LoadProviders=output
		
		Set manager=Nothing
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function LoadProvidersByPatientID(patientID, providerID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CProviderManager		
		manager.SelectProvidersByPatientID patientID
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
	    output="<select name='ProviderID' id='ProviderID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ProviderID = CInt(providerID), " selected ", "")
		    output=output & "<option value='" & value.ProviderID & "'" & selected & ">" & value.FirstName & " " & IIf(value.Mi <> "", value.Mi & " ", " ") &  value.LastName & "</option>"
	    Next
    	
	    output=output & "</select>"
				
		LoadProvidersByPatientID=output
		
		Set manager=Nothing
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function LoadProvidersByPatientIDWithOnChange(patientID, providerID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CProviderManager		
		manager.SelectProvidersByPatientID patientID
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
	    output="<select name='ProviderID' id='ProviderID' onChange='javascript:submitform();'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ProviderID = CInt(providerID), " selected ", "")
		    output=output & "<option value='" & value.ProviderID & "'" & selected & ">" & value.FirstName & " " & IIf(value.Mi <> "", value.Mi & " ", " ") &  value.LastName & "</option>"
	    Next
    	
	    output=output & "</select>"
				
		LoadProvidersByPatientIDWithOnChange=output
		
		Set manager=Nothing
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function LoadProvidersWithOnChange(providerID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CProviderManager		
		manager.SelectProviders
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
	    output="<select name='ProviderID' id='ProviderID' onChange='javascript:submitform();'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ProviderID = CInt(providerID), " selected ", "")
		    output=output & "<option value='" & value.ProviderID & "'" & selected & ">" & value.FirstName & " " & IIf(value.Mi <> "", value.Mi & " ", " ") & value.LastName & "</option>"
	    Next
    	
	    output=output & "</select>"
				
		LoadProvidersWithOnChange=output
		
		Set manager=Nothing
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function LoadExpenseTypes(expenseTypeID)
		'On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CExpenseTypeManager		
		manager.SelectExpenseTypes
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
		output="<select name='ExpenseTypeID' id='ExpenseTypeID'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ExpenseTypeID = CInt(expenseTypeID), " selected ", "")
		    output=output & "<option value='" & value.ExpenseTypeID & "'" & selected & ">" & value.ExpenseTypeName & "</option>"
		Next

	    output=output & "</select>"
				
		LoadExpenseTypes=output
		
		Set manager=Nothing
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function LoadServices(serviceID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CServiceManager		
		manager.SelectServices
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
		output="<select name='ServiceID' id='ServiceID'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ServiceID = CInt(serviceID), " selected ", "")
		    output=output & "<option value='" & value.ServiceID & "'" & selected & ">" & value.ServiceName & "</option>"
		Next

	    output=output & "</select>"
				
		LoadServices=output
		
		Set manager=Nothing
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function LoadServicesByProviderID(providerID, serviceID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CServiceManager		
		manager.SelectServicesByProviderID providerID
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
		output="<select name='ServiceID' id='ServiceID'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ServiceID = CInt(serviceID), " selected ", "")
		    output=output & "<option value='" & value.ServiceID & "'" & selected & ">" & value.ServiceName & "</option>"
		Next

	    output=output & "</select>"
				
		LoadServicesByProviderID=output
		
		Set manager=Nothing
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function LoadServicesByProviderIDWithOnChange(providerID, serviceID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CServiceManager		
		manager.SelectServicesByProviderID providerID
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
		output="<select name='ServiceID' id='ServiceID' onChange='javascript:submitform();'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ServiceID = CInt(serviceID), " selected ", "")
		    output=output & "<option value='" & value.ServiceID & "'" & selected & ">" & value.ServiceName & "</option>"
		Next

	    output=output & "</select>"
				
		LoadServicesByProviderIDWithOnChange=output
		
		Set manager=Nothing
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function LoadServicesWithOnChange(serviceID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CServiceManager		
		manager.SelectServices
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		
		keys=list.Keys
		
		output="<select name='ServiceID' id='ServiceID' onChange='javascript:submitform();'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ServiceID = CInt(serviceID), " selected ", "")
		    output=output & "<option value='" & value.ServiceID & "'" & selected & ">" & value.ServiceName & "</option>"
		Next

	    output=output & "</select>"
				
		LoadServicesWithOnChange=output
		
		Set manager=Nothing
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function SelectClaimByID(claimID)
		Dim manager
		Dim list
		Dim value
		Dim i
		Dim keys
		
		Set manager=New CClaimManager
		manager.SelectClaimByID(claimID)
		
		Set list=manager.List
		Set manager=Nothing
		
		keys=list.Keys
		Set value=list.Item(keys(0))
		Set list=Nothing
		
		Set SelectClaimByID=value 
	End Function
	
	Public Function ViewClaims(startDate, endDate)
		'On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectClaimsByDates startDate, endDate
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewClaims = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function ViewClaimsByAmount(amount)
		'On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectClaimsByAmount amount
		AddMessage manager.Messages, true
		Set list=manager.List
		
		ViewClaimsByAmount = LoadList(list)
		
		Set manager=Nothing
		Set list=Nothing
	End Function

	Public Function ViewClaimsByMonth(startDate, endDate, claimType)
		'On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectClaimsByMonth startDate, endDate, claimType
		AddMessage manager.Messages, true
		Set list=manager.List
		
		ViewClaimsByMonth = LoadMonthlyClaims(list)
		
		Set manager=Nothing
		Set list=Nothing
	End Function

	Public Function ViewClaimsByPatientID(patientID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectClaimsByPatientID patientID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewClaimsByPatientID = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function ViewClaimsByProviderID(providerID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectClaimsByProviderID providerID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewClaimsByProviderID = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function ViewClaimsByServiceID(serviceID)
		'On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectClaimsByServiceID serviceID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewClaimsByServiceID = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function ViewClaimsByServiceIDAndDate(serviceID, startDate, endDate)
		'On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectClaimsByServiceIDAndDate serviceID, startDate, endDate
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewClaimsByServiceIDAndDate = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function ViewClaimsByClaimNumber(claimNumber)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectClaimsByClaimNumber claimNumber
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewClaimsByClaimNumber = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function ViewClaimsByInsuranceClaimNumber(claimNumber)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectClaimsByInsuranceClaimNumber claimNumber
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewClaimsByInsuranceClaimNumber = LoadList(list)
		
		Set list=Nothing
	End Function
	
	Public Function ViewClaimAmountByDates(startDate, endDate)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectClaimAmountByDates startDate, endDate
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewClaimAmountByDates = LoadClaimAmounts(list)
		
		Set list=Nothing
	End Function

	Public Function ViewFlexPayRequestAmountsByInsuranceClaimNumber(claimNumber)
		'On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectFlexPayRequestAmountsByInsuranceClaimNumber claimNumber
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewFlexPayRequestAmountsByInsuranceClaimNumber = LoadFlexPayRequestAmounts(list)
		
		Set list=Nothing
	End Function

	Public Function ViewFlexPayRequestAmountByDates(startDate, endDate)
		'On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectFlexPayRequestAmountByDates startDate, endDate
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewFlexPayRequestAmountByDates = LoadFlexPayRequestAmounts(list)
		
		Set list=Nothing
	End Function
	
	Public Function ViewPaidClaimAmounts(startDate, endDate)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectPaidClaimAmountsByDate startDate, endDate
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewPaidClaimAmounts = LoadPaidClaimAmounts(list)
		
		Set list=Nothing
 	End Function
	
	Public Function ViewUnpaidClaims(startDate, endDate)
		'On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectUnpaidClaimsByDates startDate, endDate
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewUnpaidClaims = LoadList(list)
		
		Set list=Nothing
	End Function
    
    Public Function ViewClaimLineItemsByClaimID(claimID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CClaimManager
		manager.SelectClaimLineItemsByClaimID claimID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewUnpaidClaims = LoadLineItemList(list)
		
		Set list=Nothing
    End Function
	
	Public Function Delete(value)
		Dim manager
		
		Set manager=New CClaimManager
		Set manager.Claim=value
		
		Delete=manager.Delete()
		Set manager=Nothing
	End Function
	
	Public Function Save(value)
		Dim manager
		
		Set manager=New CClaimManager
		Set manager.Claim=value
		
		Save=manager.Save()
		Set manager=Nothing
	End Function
	
	Public Function Update(value)
		Dim manager
		
		Set manager=New CClaimManager
		Set manager.Claim=value
		
		Update=manager.Update()
		Set manager=Nothing
	End Function
	
End Class
%>