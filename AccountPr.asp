<!-- #include file="App_Code/AccountPageManager.asp" -->
<%
Dim processType
Dim manager
Dim value
Dim url

processType=Request.Form("ProcessType")
Set value=New CAccount

value.AccountID=Request.Form("AccountID")
value.PayeeID = Request.Form("PayeeID")
value.AccountNumber = Request.Form("AccountNumber")

Set manager=New CAccountPageManager

If processType="Add" Then
	If manager.Save(value) Then
		url="AccountsView.asp"
	Else
		url="AddAccount.asp?saved=0&error=" & manager.Messages
	End If
ElseIf processType="Update" Then
	If manager.Update(value) Then
		url="AccountsView.asp"
	Else
		url="EditAccount.asp?saved=0&error=" & manager.Messages
	End If
Else
End If

Set manager=Nothing
Set value=Nothing
Response.Redirect(url)

%>