<!-- #include file="App_Code/InvoicePageManager.asp" -->
<%
Dim message
Dim manager
Dim value
Dim url

Set value=New CInvoicePayment
Set manager=New CInvoicePageManager

value.PaymentID=Request.QueryString("PaymentID")
value.InvoiceID = Request.QueryString("InvoiceID")

If manager.DeletePayment(value) Then
	url="EditInvoice.asp?InvoiceID=" & value.InvoiceID
Else
	url=""
	message = manager.Message
End If

Set manager=Nothing
Set value=Nothing
If url <> "" Then
Response.Redirect(url)
End If
%>
<html>
	<head>
		<title>Delete Invoice Payment</title>
	</head>
	<body>
		<h5>Delete Invoice Payment</h5>
		<%= message %>
	</body>
</html>