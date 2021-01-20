<!-- #include file="InvoiceManager.asp" -->
<!-- #include file="Invoice.asp" -->
<!-- #include file="InvoicePaymentManager.asp" -->
<!-- #include file="InvoicePayment.asp" -->
<!-- #include file="ProviderManager.asp" -->
<!-- #include file="Provider.asp" -->
<!-- #include file="AccountManager.asp" -->
<!-- #include file="Account.asp" -->
<!-- #include file="ClaimManager.asp" -->
<!-- #include file="Claim.asp" -->
<!-- #include file="ExpenseType.asp" -->
<!-- #include file="Patient.asp" -->
<!-- #include file="Service.asp" -->
<!-- #include file="Medication.asp" -->
<!-- #include file="PayeeManager.asp" -->
<!-- #include file="Payee.asp" -->
<!-- #include file="Common.asp" -->
<%
Class CInvoicePageManager
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
		Dim amountTotal
		Dim paidTotal
		
		amountTotal = 0
		paidTotal = 0
		keys=list.Keys
		
		output="<table class='withborder' cellpadding='1' cellspacing='1'>"
		output=output & "<tr>"
		output=output & "<th>Payee</th>"
		output=output & "<th>Account Number</th>"
		output=output & "<th>Invoice Number</th>"
		output=output & "<th>Amount</th>"
		output=output & "<th>Due Date</th>"
		output=output & "<th>Amount Paid</th>"
		output=output & "<th></th>"
		output=output & "<th></th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			cls = IIf(cls="rowalt", "row", "rowalt")
			
			Set value=list.Item(keys(i))
			amountTotal = amountTotal + value.Amount
			paidTotal = paidTotal + value.AmountPaid
			
			output=output & "<tr class='" & cls & "'>"
			output=output & "<td>" & value.PayeeName & "</td>"
			output=output & "<td>" & value.AccountNumber & "</td>"
			output=output & "<td>"
			output=output & IIf(value.InvoiceNumber = "", value.InvoiceNumber, "<a href=javascript:popUpBig('Invoices/" & value.InvoiceNumber & ".pdf');>" & value.InvoiceNumber & "</a>")
			output=output & "</td>"
			output=output & "<td align='right'>" & FormatNumber(value.Amount, 2) & "</td>"
			output=output & "<td>" & value.DueDate & "</td>"
			'output=output & "<td>" & IIf(value.DatePaid = "1/1/1900", "", value.DatePaid) & "</td>"
			output=output & "<td align='right'>" & IIf(CDbl(value.AmountPaid) = 0.00, "", FormatNumber(value.AmountPaid, 2)) & "</td>"
			output=output & "<td>"
			output=output & "<a href='EditInvoice.asp?InvoiceID=" & value.InvoiceID & "' class='image' title=''><img alt='Edit' src='images/pencil.svg' width='16' height='16' border='0' /></a>"			
			output=output & "</td>"
			output=output & "<td>"
			output=output & "<a href='EditClaim.asp?ClaimID=" & value.ClaimID & "' class='image' title=''><img alt='View Claim' src='images/pencil-fill.svg' width='16' height='16' border='0' /></a>"
			output=output & "</td>"
			output=output & "</tr>"
		Next

		output=output & "<tr>"
		output=output & "<th></th>"
		output=output & "<th></th>"
		output=output & "<th></th>"
		output=output & "<th class='align_right'>" & FormatNumber(amountTotal, 2) & "</th>"
		output=output & "<th></th>"
		output=output & "<th class='align_right'>" & FormatNumber(paidTotal, 2) & " </th>"
		output=output & "<th></th>"
		output=output & "<th></th>"
		output=output & "</tr>"

		Set value=Nothing
		
		output=output & "</table>"
		
		LoadList = output
	End Function
	
	Private Function LoadInvoicesToPayList(list)
		Dim i
		Dim cls
		Dim keys
		Dim value
		Dim output
		Dim amountTotal
		
		amountTotal = 0
		keys=list.Keys
		
		output="<table class='withborder' cellpadding='1' cellspacing='1'>"
		output=output & "<tr>"
		output=output & "<th>Payee</th>"
		output=output & "<th>Amount</th>"
		output=output & "<th></th>"
	
		For i=0 To list.Count - 1
			cls = IIf(cls="rowalt", "row", "rowalt")
			
			Set value=list.Item(keys(i))
			amountTotal = amountTotal + value.Amount
			output=output & "<tr class='" & cls & "'>"
			output=output & "<td>" & value.PayeeName & "</td>"
			output=output & "<td align='right'>" & FormatNumber(value.Amount, 2) & "</td>"
			output=output & "<td>"
			output=output & "<a href='InvoicesSearchView.asp?PayeeID=" & value.PayeeID & "&searchtype=rbPayeeID' class='image' title=''><img alt='View Invoices' src='images/binoculars.svg' width='16' height='16' border='0' /></a>"
			output=output & "</td>"
			output=output & "</tr>"
		Next
		
		output=output & "<tr>"
		output=output & "<th></th>"
		output=output & "<th align='right'>" & FormatNumber(amountTotal, 2) & "</th>"
		output=output & "<th></th>"
		output=output & "</tr>"
		output=output & "</table>"
		
		LoadInvoicesToPayList = output
		Set value=Nothing
	End Function
	
	Private Function LoadPaymentList(list)
		Dim i
		Dim cls
		Dim keys
		Dim value
		Dim output
		Dim total
		
		total = 0
		keys=list.Keys
		
		output="<table class='withborder' cellpadding='1' cellspacing='1'>"
		output=output & "<tr>"
		output=output & "<th>Date Paid</th>"
		output=output & "<th>Amount Paid</th>"
		output=output & "<th></th>"
		output=output & "<th></th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			cls = IIf(cls="rowalt", "row", "rowalt")
			
			Set value=list.Item(keys(i))
			total = total + value.Amount
			
			output=output & "<tr class='" & cls & "'>"
			output=output & "<td>" & value.DatePaid & "</td>"
			output=output & "<td align='right'>" & FormatNumber(value.Amount, 2) & "</td>"
			output=output & "<td>"
			output=output & "<a href='EditInvoicePayment.asp?PaymentID=" & value.PaymentID & "' class='image' title=''><img alt='Edit' src='images/pencil.svg' width='16' height='16' border='0' /></a>"
			output=output & "</td>"
			output=output & "<td>"
			output=output & "<a onclick='return confirmSubmit()' href='DeleteInvoicePaymentPr.asp?PaymentID=" & value.PaymentID & "&InvoiceID=" & value.InvoiceID & "' class='image' title=''><img alt='Delete' src='images/x.svg' width='16' height='16' border='0' /></a>"
			output=output & "</td>"
			output=output & "</tr>"
		Next
		
		output=output & "<tr>"
		output=output & "<th>Total</th>"
		output=output & "<th align='right'>" & FormatNumber(total, 2) & "</th>"
		output=output & "<th></th>"
		output=output & "<th></th>"
		output=output & "</tr>"

		Set value=Nothing
		
		output=output & "</table>"
		
		LoadPaymentList = output
	End Function
	
	Public Property Get Messages()
		Messages=mMessages
	End Property

	Public Function IIf(expression, trueValue, falseValue)
		If expression Then
			IIf = trueValue
		Else
			IIf = falseValue
		End If
	End Function

	Public Function LoadInvoiceNumbersWithOnFocus(invoiceNumber)
		'On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CInvoiceManager		
		manager.SelectInvoiceNumbers
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='InvoiceNumber' id='InvoiceNumber' onFocus='setRadioIndex(document.form.searchtype, &quot;rbInvoiceNumber&quot;);'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.invoiceNumber = invoiceNumber, " selected ", "")
		    output=output & "<option value='" & value.invoiceNumber & "'" & selected & ">" & value.InvoiceNumber & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadInvoiceNumbersWithOnFocus=output
		
		Set value = Nothing
		Set list = Nothing
	End Function
	
	Public Function LoadInvoices(invoiceID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CInvoiceManager		
		manager.SelectInvoices
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='InvoiceID' id='InvoiceID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.InvoiceID = CInt(invoiceID), " selected ", "")
		    output=output & "<option value='" & value.InvoiceID & "'" & selected & ">" & value.InvoiceNumber & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadInvoices=output
		
		Set value = Nothing
		Set list = Nothing
	End Function
	
	' Public Function LoadCategories(categoryID)
		' On Error Resume Next
		' Dim selected
		' Dim manager
		' Dim output
		' Dim value
		' Dim list
		' Dim keys
		' Dim i
		
		' Set manager=New CCategoryManager		
		' manager.SelectCategories
		
		' Set list=manager.List
		' AddMessage "list.Count=" & list.Count, false
		' Set manager=Nothing
		
		' keys=list.Keys
		
		' output="<select name='CategoryID'>"
		' output=output & "<option value='0'>Select</option>"
		' For i=0 To list.Count -1 
		    ' Set value=list.Item(keys(i))
			' selected = IIf(value.CategoryID = CInt(categoryID), " selected ", "")
		    ' output=output & "<option value='" & value.CategoryID & "'" & selected & ">" & value.CategoryName & "</option>"
		' Next

	    ' output=output & "</select>"
				
		' LoadCategories=output
		
		' Set list=Nothing
		' Set value=Nothing
	' End Function
	
	' Public Function LoadCategoriesWithOnChange(categoryID)
		' On Error Resume Next
		' Dim selected
		' Dim manager
		' Dim output
		' Dim value
		' Dim list
		' Dim keys
		' Dim cls
		' Dim i
		
		' Set manager=New CCategoryManager		
		' manager.SelectCategories
		
		' Set list=manager.List
		' AddMessage "list.Count=" & list.Count, false
		' Set manager=Nothing
		
		' keys=list.Keys
		
	    ' output="<select name='CategoryID' onChange='javascript:submitform();'>"
		' output=output & "<option value='0'>Select</option>"
	    ' For i=0 To list.Count -1 
		    ' Set value=list.Item(keys(i))
			' selected = IIf(value.CategoryID = CInt(categoryID), " selected ", "")
		    ' output=output & "<option value='" & value.CategoryID & "'" & selected & ">" & value.CategoryName & "</option>"
	    ' Next
    	
	    ' output=output & "</select>"
		' Set list=Nothing
		' Set value=Nothing
				
		' LoadCategoriesWithOnChange=output
		
		' Set value = Nothing
		' Set list = Nothing
		' Set manager = Nothing
	' End Function
	
	' Public Function LoadPaychecks(paycheckID)
		' On Error Resume Next
		' Dim selected
		' Dim manager
		' Dim output
		' Dim value
		' Dim list
		' Dim keys
		' Dim cls
		' Dim i
		
		' Set manager=New CPaycheckManager		
		' manager.SelectPaychecks
		
		' Set list=manager.List
		' AddMessage "list.Count=" & list.Count, false
		' Set manager=Nothing
		
		' keys=list.Keys
		
	    ' output="<select name='PaycheckID'>"
		' output=output & "<option value='0'>Select</option>"
	    ' For i=0 To list.Count -1 
		    ' Set value=list.Item(keys(i))
			' selected = IIf(value.PaycheckID = CInt(paycheckID), " selected ", "")
		    ' output=output & "<option value='" & value.PaycheckID & "'" & selected & ">" & value.PayDateName & "</option>"
	    ' Next
    	
	    ' output=output & "</select>"
		' Set list=Nothing
		' Set value=Nothing
				
		' LoadPaychecks=output
		
		' Set value = Nothing
		' Set list = Nothing
		' Set manager = Nothing
	' End Function
	
	Public Function LoadClaimsByProviderID(providerID, claimID)
		'On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CClaimManager		
		manager.SelectClaimsByProviderID providerID
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='ClaimID' id='ClaimID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ClaimID = CInt(claimID), " selected ", "")
		    output=output & "<option value='" & value.ClaimID & "'" & selected & ">" & value.ClaimNumber & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadClaimsByProviderID=output
		
		Set value = Nothing
		Set list = Nothing
		Set manager = Nothing
	End Function
	
	Public Function LoadClaimsByPayeeID(payeeID, claimID)
		'On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CClaimManager		
		manager.SelectClaimsByPayeeID payeeID
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='ClaimID' id='ClaimID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(CInt(value.ClaimID) = CInt(claimID), " selected ", "")
		    output=output & "<option value='" & value.ClaimID & "'" & selected & ">" & value.ClaimNumber & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadClaimsByPayeeID=output
		
		Set value = Nothing
		Set list = Nothing
		Set manager = Nothing
	End Function
	
	Public Function LoadClaimsWithoutInvoiceByPayeeID(payeeID, claimID)
		'On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CClaimManager		
		manager.SelectClaimsWithoutInvoiceByPayeeID payeeID
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='ClaimID' id='ClaimID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(CInt(value.ClaimID) = CInt(claimID), " selected ", "")
		    output=output & "<option value='" & value.ClaimID & "'" & selected & ">" & value.ClaimNumber & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadClaimsWithoutInvoiceByPayeeID=output
		
		Set value = Nothing
		Set list = Nothing
		Set manager = Nothing
	End Function
	
	' Public Function LoadPaychecksWithOnChange(paycheckID)
		' On Error Resume Next
		' Dim selected
		' Dim manager
		' Dim output
		' Dim value
		' Dim list
		' Dim keys
		' Dim cls
		' Dim i
		
		' Set manager=New CPaycheckManager		
		' manager.SelectPaychecks
		
		' Set list=manager.List
		' AddMessage "list.Count=" & list.Count, false
		' Set manager=Nothing
		
		' keys=list.Keys
		
	    ' output="<select name='PaycheckID' onChange='javascript:submitform();'>"
		' output=output & "<option value='0'>Select</option>"
	    ' For i=0 To list.Count -1 
		    ' Set value=list.Item(keys(i))
			' selected = IIf(value.PaycheckID = CInt(paycheckID), " selected ", "")
		    ' output=output & "<option value='" & value.PaycheckID & "'" & selected & ">" & value.PayDateName & "</option>"
	    ' Next
    	
	    ' output=output & "</select>"
		' Set list=Nothing
		' Set value=Nothing
				
		' LoadPaychecksWithOnChange=output
		
		' Set value = Nothing
		' Set list = Nothing
		' Set manager = Nothing
	' End Function
	
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
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='PayeeID' id='PayeeID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.PayeeID = CInt(payeeID), " selected ", "")
		    output=output & "<option value='" & value.PayeeID & "'" & selected & ">" & value.PayeeName & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadPayees=output
		
		Set value = Nothing
		Set list = Nothing
		Set manager = Nothing
	End Function
	
	Public Function LoadPayeesWithOnChange(payeeID)
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
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='PayeeID' id='PayeeID' onChange='javascript:submitform();'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.PayeeID = CInt(payeeID), " selected ", "")
		    output=output & "<option value='" & value.PayeeID & "'" & selected & ">" & value.PayeeName & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadPayeesWithOnChange=output
		
		Set value = Nothing
		Set list = Nothing
		Set manager = Nothing
	End Function
	
	Public Function LoadPayeesWithOnFocus(payeeID)
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
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='PayeeID' id='PayeeID' onFocus='setRadioIndex(document.form.searchtype, &quot;rbPayeeID&quot;);'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.PayeeID = CInt(payeeID), " selected ", "")
		    output=output & "<option value='" & value.PayeeID & "'" & selected & ">" & value.PayeeName & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadPayeesWithOnFocus=output
		
		Set value = Nothing
		Set list = Nothing
		Set manager = Nothing
	End Function
	
	Public Function LoadAccountsByPayeeID(payeeID, accountID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CAccountManager		
		manager.SelectAccountsByPayeeID payeeID
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='AccountID' id='AccountID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.AccountID = CInt(accountID), " selected ", "")
		    output=output & "<option value='" & value.AccountID & "'" & selected & ">" & value.AccountNumber & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadAccountsByPayeeID=output
		
		Set value = Nothing
		Set list = Nothing
		Set manager = Nothing
	End Function
	
	' Public Function LoadAccountsByProviderID(providerID)
		' On Error Resume Next
		' Dim selected
		' Dim manager
		' Dim output
		' Dim value
		' Dim list
		' Dim keys
		' Dim i
		
		' Set manager=New CAccountManager		
		' manager.SelectAccountsByProviderID providerID
		
		' Set list=manager.List
		' AddMessage "list.Count=" & list.Count, false
		' Set manager=Nothing
		
		' keys=list.Keys
		
	    ' output="<select name='AccountID'>"
		' output=output & "<option value='0'>Select</option>"
	    ' For i=0 To list.Count -1 
		    ' Set value=list.Item(keys(i))
			' selected = IIf(value.providerID = CInt(providerID), " selected ", "")
		    ' output=output & "<option value='" & value.AccountID & "'" & selected & ">" & value.AccountNumber & "</option>"
	    ' Next
    	
	    ' output=output & "</select>"
		' Set list=Nothing
		' Set value=Nothing
				
		' LoadAccountsByProviderID=output
		
		' Set value = Nothing
		' Set list = Nothing
		' Set manager = Nothing
	' End Function
	
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
		Set manager=Nothing
		
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
		Set manager=Nothing
		
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
		
		Set list=Nothing
		Set value=Nothing
	End Function

	Public Function SelectInvoiceByID(invoiceID)
		Dim manager
		Dim list
		Dim value
		Dim i
		Dim keys
		
		Set manager=New CInvoiceManager
		manager.SelectInvoiceByID invoiceID
		
		Set list=manager.List
		Set manager=Nothing
		
		keys=list.Keys
		Set value=list.Item(keys(0))
		Set list=Nothing
		
		Set SelectInvoiceByID=value 
	End Function
	
	Public Function ViewInvoices()
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CInvoiceManager		
		manager.SelectInvoices
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewInvoices = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function ViewInvoicesByAccountNumber(accountNumber)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CInvoiceManager		
		manager.SelectInvoicesByAccountNumber accountNumber
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewInvoicesByAccountNumber = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function ViewInvoicesByClaimNumber(claimNumber)
		'On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CInvoiceManager
		manager.SelectInvoicesByClaimNumber claimNumber
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewInvoicesByClaimNumber = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function ViewInvoicesByInvoiceNumber(invoiceNumber)
		'On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CInvoiceManager
		manager.SelectInvoicesByInvoiceNumber invoiceNumber
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewInvoicesByInvoiceNumber = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function ViewInvoicesByServiceDate(startDate, endDate)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CInvoiceManager
		manager.SelectInvoicesByServiceDate startDate, endDate
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewInvoicesByServiceDate = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function ViewInvoicesByPayeeID(payeeID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CInvoiceManager		
		manager.SelectInvoicesByPayeeID payeeID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewInvoicesByPayeeID = LoadList(list)
		
		Set list=Nothing
	End Function

	' Public Function ViewInvoicesByProviderID(providerID)
		' On Error Resume Next
		
		' Dim manager
		' Dim list
		
		' Set manager=New CInvoiceManager		
		' manager.SelectInvoicesByProviderID providerID
		
		' Set list=manager.List
		' Set manager=Nothing
		
		' ViewInvoicesByProviderID = LoadList(list)
		
		' Set list=Nothing
	' End Function

	' Public Function ViewInvoicesByCategoryID(categoryID)
		' On Error Resume Next
		
		' Dim manager
		' Dim list
		
		' Set manager=New CInvoiceManager		
		' manager.SelectInvoicesByCateogryID categoryID
		
		' Set list=manager.List
		' Set manager=Nothing
		
		' ViewInvoicesByCategoryID = LoadList(list)
		
		' Set list=Nothing
	' End Function

	Public Function ViewInvoicesToPay()
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CInvoiceManager		
		manager.SelectInvoiceAmountsToPay
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewInvoicesToPay = LoadInvoicesToPayList(list)
		
		Set list=Nothing
	End Function
	
	Public Function ViewInvoicePaymentsByInvoiceID(invoiceID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CInvoicePaymentManager		
		manager.SelectInvoicePaymentsByInvoiceID invoiceID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewInvoicePaymentsByInvoiceID = LoadPaymentList(list)
		
		Set list=Nothing
	End Function

	Public Function Save(value)
		Dim manager
		
		Set manager=New CInvoiceManager
		Set manager.Invoice=value
		
		If manager.Save() Then
			Save=true
		Else
			Save=false
			AddMessage manager.Messages, true
		End If
		
		Set manager=Nothing
	End Function
	
	Public Function Update(value)
		Dim manager
		
		Set manager=New CInvoiceManager
		Set manager.Invoice=value
		
		If manager.Update() Then
			Update=true
		Else
			Update=false
			AddMessage manager.Messages, true
		End If
		
		Set manager=Nothing
	End Function
	
	Public Function DeletePayment(value)
		Dim manager
		
		Set manager=New CInvoicePaymentManager
		Set manager.InvoicePayment=value
		
		If manager.Delete() Then
			DeletePayment=true
		Else
			DeletePayment=false
			AddMessage manager.Messages, true
		End If
		
		Set manager=Nothing
	End Function
End Class
%>