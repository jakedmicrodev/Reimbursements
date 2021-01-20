<!-- #include file="App_Code/ClaimPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim claimID
Dim patientList
Dim providerList
Dim claimLineItemList

claimID = Request.QueryString("ClaimID")

Set manager = New CClaimPageManager
Set value = manager.SelectClaimByID(claimID)
providerList = manager.LoadProviders(value.ProviderID)
patientList = manager.LoadPatients(value.PatientID)
'claimLineItemList = manager.ViewClaimLineItemsByClaimID(claimID)
	
Set manager = Nothing
%>
<html>
	<head>
		<title>Edit Invoice</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/table.css" type="text/css" />
		<link rel="stylesheet" type="text/css" media="all" href="jscalendar-1.0/skins/aqua/theme.css" title="Aqua" />
		<script type="text/javascript" src="jscalendar-1.0/calendar.js"></script>
		<script type="text/javascript" src="jscalendar-1.0/lang/calendar-en.js"></script>
		<script type="text/javascript" src="jscalendar-1.0/calendar-setup.js"></script>
		<script language="JavaScript">
			function popUp(URL) {
			day = new Date();
			id = day.getTime();
			eval("page" + id + " = window.open(URL, '" + id + "', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=yes,width=300,height=300,left = 200,top = 100');");
			}
		</script>
	</head>
	<body>
	<!-- #include file="menu\menu.inc" -->
		<form  action="ClaimPr.asp" method="post" name="form">
			<input type="hidden" name="ProcessType" value="Update">
			<input type="hidden" name="ClaimID" value="<%= claimID %>">			
			<table>
				<tr>
					<th colspan="2">Add Claim</th>
				<tr>
				<tr>
					<td>Patient</td>
					<td><%= patientList %></td>
				</tr>
				<tr>
					<td>Provider</td>
					<td><%= providerList %></td>
				</tr>
				<tr>
					<td>Patient Account Number</td>
					<td><input name="AccountNumber" size="15" value="<%= value.AccountNumber %>"></td>
				</tr>
				<tr>
					<td>Claim Number</td>
					<td><input name="ClaimNumber" size="15" value="<%= value.ClaimNumber %>"></td>
				</tr>
				<tr>
					<td class="rowalt" colspan="2" align="left">
                        <input type="submit" value=":: Save ::" />&nbsp;
                        <input type="button" onclick="javascript:popUp('AddClaimLineItem.asp?ClaimID=<%= value.ClaimID %>')" value="Add Line Item"/></td>
				</tr>
				<tr><td colspan="2">&nbsp;</td></tr>
			</table>
            <table width="600">
				<tr>
					<th colspan="2">Line Items</th>
				</tr>
				<tr>
					<td colspan="2"><%= claimLineItemList%></td>
				</tr>
            </table>
		</form>
	</body>
</html>
<%
Set value = Nothing
%>