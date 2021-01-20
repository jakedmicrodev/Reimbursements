<!-- #include file="App_Code/GiftPageManager.asp" -->
<%
Dim processType
Dim manager
Dim value
Dim url

processType=Request.Form("ProcessType")
Set value=New CGift

value.MemberID=Request.Form("MemberID")
value.OccasionID=Request.Form("OccasionID")
value.GiftName=Request.Form("GiftName")
value.Location=Request.Form("Location")
value.Cost=Request.Form("Cost")

If processType="Update" Then
	value.GiftID=Request.Form("GiftID")
End If

Set manager=New CGiftPageManager

If processType="Add" Then
	If manager.Save(value) Then
		url="ViewGifts.asp"
	Else
		url="AddGift.asp?saved=0&error=" & manager.Messages
	End If
ElseIf processType="Update" Then
	If manager.Update(value) Then
		url="ViewGifts.asp"
	Else
		url="EditGift.asp?saved=0&error=" & manager.Messages
	End If
Else
End If

Set manager=Nothing
Set value=Nothing
Response.Redirect(url)

%>