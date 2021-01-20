<!-- #include file="App_Code/GiftPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim value
Dim manager
Dim roomList
Dim memberList
Dim memberID
Dim occasionID
Dim occasionList

giftID = Request.QueryString("giftID")

Set manager = New CGiftPageManager
Set value = manager.SelectGiftByID(giftID)

memberList = manager.LoadMembers(value.MemberID)
occasionList = manager.LoadOccasions(value.OccasionID)

Set manager = Nothing
%>
<html>
	<head>
		<title>Add Gift</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
	</head>
	<body>
	<!-- #include file="menu\menu.inc" -->
		<form  action="GiftPr.asp" method="post" name="form">
			<input type="hidden" name="ProcessType" value="Update">
			<input type="hidden" name="GiftID" value="<%= giftID %>">
			<table>
				<tr>
					<th colspan="2">Edit Gift</th>
				<tr>
				<tr>
					<td>Requestor</td>
					<td><%= memberList %></td>
				</tr>
				<tr>
					<td>Occasion</td>
					<td><%= occasionList %></td>
				</tr>
				<tr>
					<td>Gift</td>
					<td><input name="GiftName" size="30" value="<%= value.GiftName %>"></td>
				</tr>
				<tr>
					<td>Where to Find</td>
					<td><input name="Location" size="30" value="<%= value.Location %>"></td>
				</tr>
				<tr>
					<td>Cost</td>
					<td><input name="Cost" size="5" value="<%= value.Cost %>"></td>
				</tr>
				<tr>
					<td class="rowalt" colspan="2" align="left"><input type="submit" value=":: Save ::" /></td>
				</tr>
			</table>
			<%
			If saved=1 Then
				Response.Write("<font color=green><b>Gift Information Saved</b></font>")
			End If
			%>
		</form>
	</body>
</html>