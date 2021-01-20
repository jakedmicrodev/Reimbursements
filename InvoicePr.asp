<!-- #include file="App_Code/InvoicePageManager.asp" -->
<%
Dim processType
Dim manager
Dim value
Dim url

processType=Request.Form("ProcessType")
Set value=New CInvoice

value.InvoiceID=Request.Form("InvoiceID")
value.PayeeID = Request.Form("PayeeID")
value.AccountID = Request.Form("AccountID")
value.InvoiceNumber = Request.Form("InvoiceNumber")
value.Amount = Request.Form("Amount")
value.DueDate = Request.Form("DueDate")
value.ClaimID = Request.Form("ClaimID")

Set manager=New CInvoicePageManager

If processType="Add" Then
	If manager.Save(value) Then
		url="InvoicesView.asp?PayeeID=" & value.PayeeID
	Else
		url="AddInvoice.asp?saved=0&error=" & manager.Messages
	End If
ElseIf processType="Update" Then
	If manager.Update(value) Then
		url="InvoicesView.asp?PayeeID=" & value.PayeeID
	Else
		url="EditInvoice.asp?saved=0&error=" & manager.Messages
	End If
Else
End If

Set manager=Nothing
Set value=Nothing
Response.Redirect(url)

%>
<html>
	<head>
	</head>
	<body>
	<%= manager.Messages %>
	</body>
</html>