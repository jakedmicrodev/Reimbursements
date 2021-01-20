<!-- #include file="App_Code/ClaimPageManager.asp" -->
<%
Dim processType
Dim manager
Dim value
Dim url

processType=Request.Form("ProcessType")
Set value=New CClaim
Set manager=New CClaimPageManager

value.ClaimID=Request.Form("ClaimID")
value.PatientID=Request.Form("PatientID")
value.ProviderID=Request.Form("ProviderID")
value.PayeeID=Request.Form("PayeeID")
value.ServiceID=Request.Form("ServiceID")
value.ExpenseDate=Request.Form("ExpenseDate")
value.ExpenseAmount=Request.Form("ExpenseAmount")
value.ExpenseTypeID=Request.Form("ExpenseTypeID")
value.ClaimNumber=Request.Form("ClaimNumber")
value.InsuranceClaimNumber=Request.Form("InsuranceClaimNumber")
value.PaidCD=manager.IIf(Request.Form("PaidCD") <> "", 1, 0)
value.MedicationID=manager.IIf(Request.Form("MedicationID") = "", 0 , Request.Form("MedicationID"))
value.MedicationAmount=manager.IIf(Request.Form("MedicationAmount") = "", 0 , Request.Form("MedicationAmount"))

' Response.Write("ClaimID " & value.ClaimID & "<br/>")
' Response.Write("PatientID " & value.PatientID & "<br/>")
' Response.Write("ProviderID " & value.ProviderID & "<br/>")
' Response.Write("PayeeID " & value.PayeeID & "<br/>")
' Response.Write("ServiceID " & value.ServiceID & "<br/>")
' Response.Write("ExpenseDate " & value.ExpenseDate & "<br/>")
' Response.Write("ExpenseAmount " & value.ExpenseAmount & "<br/>")
' Response.Write("ExpenseTypeID " & value.ExpenseTypeID & "<br/>")
' Response.Write("ClaimNumber " & value.ClaimNumber & "<br/>")
' Response.Write("InsuranceClaimNumber " & value.InsuranceClaimNumber & "<br/>")
' Response.Write("PaidCD " & value.PaidCD & "<br/>")
' Response.Write("MedicationID " & value.MedicationID & "<br/>")
' Response.Write("MedicationAmount " & value.MedicationAmount & "<br/>")
' Response.Write("Hello")

If processType="Add" Then
	If manager.Save(value) Then
		url="ClaimsSearch.asp"
	Else
		url="AddClaim.asp?saved=0&error=" & manager.Messages
	End If
ElseIf processType="Update" Then
	If manager.Update(value) Then
		url="ClaimsSearch.asp"
	Else
		url="EditClaim.asp?saved=0&error=" & manager.Messages
	End If
Else
End If

Set manager=Nothing
Set value=Nothing
Response.Redirect(url)

%>