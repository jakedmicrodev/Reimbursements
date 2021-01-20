<%
Class CInvoicePaymentManager
	Private mInvoicePayment
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
		AddMessage "State: " & mConnection.State, true
		Set mAccount=Nothing
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

    Sub SetCustomError(strMessage)
        'Display custom message and information from VBScript Err object.

        AddMessage "<br/>" & strMessage & "<br/>" & _
          "Number (dec) : " & Err.Number & "<br/>" & _
          "Number (hex) : &H" & Hex(Err.Number) & "<br/>" & _
          "Description  : " & Err.Description & "<br/>" & _
          "Source       : " & Err.Source, true
        Err.Clear
    End Sub

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
	
	Private Function FillInvoicePayment(rs)
		Dim value
		Set value=New CInvoicePayment

		value.PaymentID=GetFieldValue(rs,"PaymentID","Numeric")
		value.InvoiceID=GetFieldValue(rs,"InvoiceID","Numeric")
		value.Amount=GetFieldValue(rs,"Amount","Numeric")
		value.DatePaid=GetFieldValue(rs,"DatePaid","Text")
		value.PayeeName=GetFieldValue(rs,"PayeeName","Text")
			
		Set FillInvoicePayment = value
	End Function

	'Public Properties
	Public Property Set InvoicePayment(value)
		Set mInvoicePayment=value
	End Property
	
	Public Property Get List()
		Set List=mList
	End Property
	
	Public Property Get Messages()
		Messages=mMessages
	End Property
	
	'Public Methods
	Public Sub SelectInvoicePaymentByID(paymentID)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInvoicePaymentByID " & _
			AddArgument(paymentID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Account by ID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInvoicePayment(rs)
			mList.Add CStr(value.PaymentID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
		
	Public Sub SelectInvoicePaymentsByInvoiceID(invoiceID)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelInvoicePaymentsByInvoiceID " & _
			AddArgument(invoiceID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Payment by InvoiceID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillInvoicePayment(rs)
			mList.Add CStr(value.PaymentID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
	
	Public Function Delete()
		On Error Resume Next
		Dim sql
		
		sql = "pDelInvoicePayment " & _
		AddArgument(mInvoicePayment.PaymentID, "NumericTerminal")
		
		mConnection.Execute sql
		CheckForError "Unable to delete Invoice Payment!"
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

		sql = "pInsInvoicePayment " & _
		AddArgument(mInvoicePayment.InvoiceID, "Numeric") & _
		AddArgument(mInvoicePayment.Amount, "Numeric") & _
		AddArgument(mInvoicePayment.DatePaid, "DateTerminal")

		mConnection.Execute sql
		CheckForError "Unable to save new Invoice Payment!"
		AddMessage sql, false
		If mErrorNumber = 0 Then
		    Save=true
		Else
		    Save=false
		End If
	End Function

	Public Function Update()
		On Error Resume Next
		Dim sql
		
		sql = "pUpdInvoicePayment " & _
		AddArgument(mInvoicePayment.PaymentID, "Numeric") & _
		AddArgument(mInvoicePayment.InvoiceID, "Numeric") & _
		AddArgument(mInvoicePayment.Amount, "Numeric") & _
		AddArgument(mInvoicePayment.DatePaid, "DateTerminal")
		
		mConnection.Execute sql
		CheckForError "Unable to update Invoice Payment!"
		If mErrorNumber = 0 Then
		    Update=true
		Else
		    Update=false
		End If
	End Function
End Class
%>