<%
Class CInvoiceManager
	Private mInvoice
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
		AddMessage "State: " & mConnection.State, false
		Set mInvoice=Nothing
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
	Private Sub AddMessage(message, add)
		If add Then
			mMessages=mMessages & message & "<br/>"
		End If
	End Sub

    ' Function StripApostrophe(text)
        ' StripApostrophe = replace (text,"'","&acute;")
    ' End Function

	' Private Function AddArgument(value, argType)
		' Dim myValue
		
	    ' Select Case argType
	        ' Case "Numeric"
	            ' myValue = value & ", "
	        ' Case "NumericTerminal"
	            ' Nothing to do just return the number
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
	
	Private Function FillInvoice(rs)
		Dim value
		Set value=New CInvoice

		value.InvoiceID=GetFieldValue(rs,"InvoiceID","Numeric")
		value.PayeeID=GetFieldValue(rs,"PayeeID","Numeric")
		value.PayeeName=GetFieldValue(rs,"PayeeName","Text")
		value.AccountID=GetFieldValue(rs,"AccountID","Numeric")
		value.AccountNumber=GetFieldValue(rs,"AccountNumber","Text")
		value.InvoiceNumber=GetFieldValue(rs,"InvoiceNumber","Text")
		value.Amount=GetFieldValue(rs,"Amount","Numeric")
		value.DueDate=GetFieldValue(rs,"DueDate","Text")
		value.AmountPaid=GetFieldValue(rs,"AmountPaid","Numeric")
		value.ClaimID=GetFieldValue(rs,"ClaimID","Numeric")
			
		Set FillInvoice = value
	End Function
	
	Private Function FillInvoiceToPay(rs)
		Dim value
		Set value=New CInvoice

		value.PayeeID=GetFieldValue(rs,"PayeeID","Numeric")
		value.PayeeName=GetFieldValue(rs,"PayeeName","Text")
		value.Amount=GetFieldValue(rs,"Amount","Numeric")
			
		Set FillInvoiceToPay = value
	End Function
	
	Private Function FillInvoiceNumber(rs)
		Dim value
		Set value=New CInvoice

		value.InvoiceNumber=GetFieldValue(rs,"InvoiceNumber","Text")
			
		Set FillInvoiceNumber = value
	End Function
	
	'Public Properties
	Public Property Set Invoice(value)
		Set mInvoice=value
	End Property
	
	Public Property Get List()
		Set List=mList
	End Property
	
	Public Property Get Messages()
		Messages=mMessages
	End Property
	
	'Public Methods
	Public Sub SelectInvoiceNumbers()
	    'On Error Resume Next
		Dim value
		Dim rs

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInvoiceNumbers", mConnection
		CheckForError "Select Invoices Failed!"

		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInvoiceNumber(rs)
			mList.Add value.InvoiceNumber, value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub

	Public Sub SelectInvoices()
	    On Error Resume Next
		Dim value
		Dim rs

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInvoices", mConnection
		CheckForError "Select Invoices Failed!"

		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInvoice(rs)
			mList.Add CStr(value.InvoiceID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub

	Public Sub SelectInvoicesByAccountNumber(accountNumber)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInvoicesByAccountNumber " & _
			AddArgument(accountNumber, "TextTerminal"), mConnection
		CheckForError "Unable to select the Invoice by Account Number!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInvoice(rs)
			mList.Add CStr(value.InvoiceID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing
	End Sub
		
	Public Sub SelectInvoicesByClaimNumber(claimNumber)
	    'On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInvoicesByClaimNumber " & _
			AddArgument(claimNumber, "TextTerminal"), mConnection
		CheckForError "Unable to select the Invoice by Claim Number!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInvoice(rs)
			mList.Add CStr(value.InvoiceID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing
	End Sub 
		
	Public Sub SelectInvoicesByInvoiceNumber(invoiceNumber)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInvoicesByInvoiceNumber " & _
			AddArgument(invoiceNumber, "TextTerminal"), mConnection
		CheckForError "Unable to select the Invoice by Invoice Number!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInvoice(rs)
			mList.Add CStr(value.InvoiceID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing
	End Sub
		
	Public Sub SelectInvoicesByServiceDate(startDate, endDate)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInvoicesByExpenseDate " & _
			AddArgument(startDate, "Date") & _
			AddArgument(endDate, "DateTerminal"), mConnection
		CheckForError "Unable to select the Invoice by Claim Number!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInvoice(rs)
			mList.Add CStr(value.InvoiceID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing
	End Sub 
		
	Public Sub SelectInvoiceByID(invoiceID)
	    'On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInvoiceByID " & _
			AddArgument(invoiceID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Invoice by ID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInvoice(rs)
			mList.Add CStr(value.InvoiceID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
		
	Public Sub SelectInvoicesByPaycheckID(paycheckID)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInvoicesByPaycheckID " & _
			AddArgument(paycheckID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Invoice by PaycheckID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInvoice(rs)
			mList.Add CStr(value.InvoiceID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
	
	Public Sub SelectInvoicesByPayeeID(payeeID)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInvoicesByPayeeID " & _
			AddArgument(payeeID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Invoice by PayeeID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInvoice(rs)
			mList.Add CStr(value.InvoiceID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
	
	Public Sub SelectInvoicesByCateogryID(categoryID)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInvoicesByCategoryID " & _
			AddArgument(categoryID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Invoice by CategoryID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInvoice(rs)
			mList.Add CStr(value.InvoiceID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
	
	Public Sub SelectInvoiceAmountsToPay()
	    On Error Resume Next
		Dim value
		Dim rs

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInvoiceAmountsToPay", mConnection
		CheckForError "Select Invoices Failed!"

		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInvoiceToPay(rs)
			mList.Add CStr(value.PayeeID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
		
	'Microsoft JET Database Engine (0x80004005) Operation must use an updateable query.
	'Modified the permissions for all users to have full control
	Public Function Save()
		On Error Resume Next
		Dim sql

		sql = "pInsInvoice " & _
		AddArgument(mInvoice.PayeeID, "Numeric") & _
		AddArgument(mInvoice.AccountID, "Numeric") & _
		AddArgument(mInvoice.InvoiceNumber, "Text") & _
		AddArgument(mInvoice.Amount, "Numeric") & _
		AddArgument(mInvoice.DueDate, "Date") & _
		AddArgument(mInvoice.ClaimID, "NumericTerminal")

		mConnection.Execute sql
		CheckForError "Unable to save new Invoice!"
		If mErrorNumber = 0 Then
		    Save=True
		Else
		    Save=False
		End If		
	End Function

	Public Function Update()
		On Error Resume Next
		Dim sql
		
		sql = "pUpdInvoice " & _
		AddArgument(mInvoice.InvoiceID, "Numeric") & _
		AddArgument(mInvoice.PayeeID, "Numeric") & _
		AddArgument(mInvoice.AccountID, "Numeric") & _
		AddArgument(mInvoice.InvoiceNumber, "Text") & _
		AddArgument(mInvoice.Amount, "Numeric") & _
		AddArgument(mInvoice.DueDate, "Date") & _
		AddArgument(mInvoice.ClaimID, "NumericTerminal")
		
		mConnection.Execute sql
		AddMessage "'" & sql & "'", false
		CheckForError "Unable to update Invoice!"
		If mErrorNumber = 0 Then
		    Update=True
		Else
		    Update=False
		End If
	End Function
End Class
%>