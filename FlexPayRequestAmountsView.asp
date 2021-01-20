<!-- #include file="App_Code/ClaimPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim view
Dim startDate
Dim endDate

startDate = Request.Form("StartDate")
endDate = Request.Form("EndDate")

Set manager=New CClaimPageManager

If startDate <> "" And endDate <> "" Then
	view=manager.ViewFlexPayRequestAmountByDates(startDate, endDate)
End If

Set manager=Nothing
%>
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>View Flex Pay Request Amounts</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" type="text/css" media="all" href="jscalendar-1.0/skins/aqua/theme.css" title="Aqua" />
		<link rel="stylesheet" href="css/table.css" type="text/css" />
		<script type="text/javascript" src="jscalendar-1.0/calendar.js"></script>
		<script type="text/javascript" src="jscalendar-1.0/lang/calendar-en.js"></script>
		<script type="text/javascript" src="jscalendar-1.0/calendar-setup.js"></script>
	</head>
	<body>
	<!-- #include file="menu\menu.inc" -->
	<form  action="FlexPayRequestAmountsView.asp" method="post" name="form">
		<table>
			<tr>
				<th colspan="2">View Flex Pay Request Amounts</th>
			</tr>
			<tr>
				<td>Start Date</td>
				<td>
					<input type="text" name="StartDate" id="StartDate" size="10" value="<%= startDate %>"> <input type="button" id="trigger1" value="..." />
					<script type="text/javascript">
					  Calendar.setup(
						{
						  inputField  : "StartDate", // ID of the input field
						  ifFormat    : "%m/%d/%Y",    // the date format
						  button      : "trigger1"      // ID of the button
						}
					  );
					</script>							
				</td>
			</tr>
			<tr>
				<td>End Date</td>
				<td>
					<input type="text" name="EndDate" id="EndDate" size="10" value="<%= endDate %>"> <input type="button" id="trigger2" value="..." />
					<script type="text/javascript">
					  Calendar.setup(
						{
						  inputField  : "EndDate", // ID of the input field
						  ifFormat    : "%m/%d/%Y",    // the date format
						  button      : "trigger2"      // ID of the button
						}
					  );
					</script>							
				</td>
			</tr>
			<tr>
				<td class="rowalt" colspan="2" align="left"><input type="submit" value=":: Submit ::" /></td>
			</tr>
		</table>
		<br/>
		<%= view %>
	</form>
	</body>
</html>