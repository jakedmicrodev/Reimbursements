<!-- #include file="App_Code/CityPageManager.asp" -->
<%
Dim processType
Dim manager
Dim value
Dim url

processType=Request.Form("ProcessType")
Set value=New CCity

value.CityName=Request.Form("CityName")

If processType="Update" Then
	value.CityID=Request.Form("CityID")
End If

Set manager=New CCityPageManager

If processType="Add" Then
	If manager.Save(value) Then
		url="CitiesView.asp"
	Else
		url="AddCity.asp?saved=0&error=" & manager.Messages
	End If
ElseIf processType="Update" Then
	If manager.Update(value) Then
		url="CitiesView.asp"
	Else
		url="EditCity.asp?saved=0&error=" & manager.Messages
	End If
Else
End If

Set manager=Nothing
Set value=Nothing
Response.Redirect(url)

%>