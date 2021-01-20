<!-- #include file="InvoicePaymentManager.asp" -->
<!-- #include file="InvoicePayment.asp" -->
<!-- #include file="Common.asp" -->
<%
Class CInvoicePaymentPageManager
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
		
		total = 0
		keys=list.Keys
		
		output="<table>"
		output=output & "<tr>"
		output=output & "<th>Payee</th>"
		output=output & "<th>Amount</th>"
		output=output & "<th>Date Paid</th>"
		output=output & "<th></th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			Set value=list.Item(keys(i))
			total = total + value.Amount
			
			output=output & "<tr>"
			output=output & "<td>" & value.PayeeName & "</td>"
			output=output & "<td>" & FormatNumber(value.Amount, 2) & "</td>"
			output=output & "<td>" & value.DatePaid & "</td>"
			output=output & "<td>"
			output=output & "<a href='EditInvoicePayment.asp?PaymentID=" & value.PaymentID & "' class='image' title=''><img alt='Edit' src='images/pencil.svg' width='16' height='16' border='0' /></a>"
			output=output & "</td>"
			output=output & "</tr>"
		Next
		
		output=output & "<tr>"
		output=output & "<th>Total</th>"
		output=output & "<th class='align_right'>" & FormatNumber(total, 2) & "</th>"
		output=output & "<th></th>"
		output=output & "<th></th>"
		output=output & "</tr>"

		Set value=Nothing
		
		output=output & "</table>"
		
		LoadList = output
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

	Public Function SelectInvoicePaymentByID(paymentID)
		On Error Resume Next
		Dim i
		Dim list
		Dim keys
		Dim value
		Dim manager
		
		Set manager=New CInvoicePaymentManager
		manager.SelectInvoicePaymentByID paymentID
		
		Set list=manager.List
		Set manager=Nothing
		
		keys=list.Keys
		Set value=list.Item(keys(0))
		Set list=Nothing
		
		Set SelectInvoicePaymentByID=value 
	End Function
	
	Public Function ViewInvoicePaymentsByInvoiceID(invoiceID)
		'On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CInvoicePaymentManager		
		manager.SelectInvoicePaymentsByInvoiceID invoiceID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewInvoicePaymentsByInvoiceID = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function Save(value)
		Dim manager
		
		Set manager=New CInvoicePaymentManager
		Set manager.InvoicePayment=value
		
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
		
		Set manager=New CInvoicePaymentManager
		Set manager.InvoicePayment=value
		
		If manager.Update() Then
			Update=true
		Else
			Update=false
			AddMessage manager.Messages, true
		End If
		
		Set manager=Nothing
	End Function
	
End Class
%>