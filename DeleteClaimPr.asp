<!-- #include file="App_Code/ClaimPageManager.asp" -->
<%
Dim message
Dim manager
Dim value
Dim url

Set value=New CClaim
Set manager=New CClaimPageManager

value.ClaimID=Request.QueryString("ClaimID")

If manager.Delete(value) Then
	url="ClaimsView.asp"
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
		<title>Delete Claim</title>
	</head>
	<body>
		<h5>Delete Claim</h5>
		<%= message %>
	</body>
</html>