<!-- #include file="App_Code/ExpenseTypePageManager.asp" -->
<%
Dim processType
Dim manager
Dim value
Dim url

processType=Request.Form("ProcessType")
Set value=New CExpenseType

value.ExpenseTypeName=Request.Form("ExpenseTypeName")

If processType="Update" Then
	value.ExpenseTypeID=Request.Form("ExpenseTypeID")
End If

Set manager=New CExpenseTypePageManager

If processType="Add" Then
	If manager.Save(value) Then
		url="ExpenseTypesView.asp"
	Else
		url="AddExpenseType.asp?saved=0&error=" & manager.Messages
	End If
ElseIf processType="Update" Then
	If manager.Update(value) Then
		url="ExpenseTypesView.asp"
	Else
		url="EditExpenseType.asp?saved=0&error=" & manager.Messages
	End If
Else
End If

Set manager=Nothing
Set value=Nothing
Response.Redirect(url)

%>