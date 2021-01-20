<!-- #include file="App_Code/MedicationPageManager.asp" -->
<%
Dim processType
Dim manager
Dim value
Dim url

processType=Request.Form("ProcessType")
Set value=New CMedication

value.MedicationName=Request.Form("MedicationName")

If processType="Update" Then
	value.MedicationID=Request.Form("MedicationID")
End If

Set manager=New CMedicationPageManager

If processType="Add" Then
	If manager.Save(value) Then
		url="MedicationsView.asp"
	Else
		url="AddMedication.asp?saved=0&error=" & manager.Messages
	End If
ElseIf processType="Update" Then
	If manager.Update(value) Then
		url="MedicationsView.asp"
	Else
		url="EditMedication.asp?saved=0&error=" & manager.Messages
	End If
Else
End If

Set manager=Nothing
Set value=Nothing
Response.Redirect(url)

%>