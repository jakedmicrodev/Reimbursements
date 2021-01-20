<!-- #include file="App_Code/InvoicePaymentPageManager.asp" -->
<%
Dim processType
Dim manager
Dim value
Dim url

processType=Request.Form("ProcessType")
Set value=New CInvoicePayment

value.InvoiceID = Request.Form("InvoiceID")
value.Amount = Request.Form("Amount")
value.DatePaid = Request.Form("DatePaid")

If processType="Update" Then
	value.PaymentID=Request.Form("PaymentID")
End If

Set manager=New CInvoicePaymentPageManager

If processType="Add" Then
	If manager.Save(value) Then
		url="UpdateConfirm.html"
	Else
		url="AddInvoicePayment.asp?saved=0&error=" & manager.Messages
	End If
ElseIf processType="Update" Then
	If manager.Update(value) Then
		'url="UpdateConfirm.html"
		url="EditInvoice.asp?InvoiceID=" & value.InvoiceID
	Else
		url="EditInvoicePayment.asp?saved=0&error=" & manager.Messages
	End If
Else
End If

Set manager=Nothing
Set value=Nothing
Response.Redirect(url)

%>