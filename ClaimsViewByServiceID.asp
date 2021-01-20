<!-- #include file="App_Code/ClaimPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim view
Dim serviceID
Dim serviceList

serviceID = Request.Form("ServiceID")
If serviceID="" Then serviceID=0

Set manager=New CClaimPageManager
serviceList = manager.LoadServicesWithOnChange(serviceID)

view=manager.ViewClaimsByServiceID(serviceID)

'Response.Write("Messages: " & manager.Messages & "<br/>")
Set manager=Nothing
%>
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>View Claims By Service</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/table.css" type="text/css" />
		<script type="text/javascript">
			function submitform() 
			{ 
				document.form.submit(); 
			}
			
			function popUp(url, w, h, t, l) {
				day = new Date();
				id = day.getTime();
				eval("page" + id + " = window.open(url, '" + id + "', 'toolbar=0,scrollbars=1,location=0,statusbar=0,menubar=0,resizable=yes,width=" + w + ",height=" + h + ",left=" + l + ",top=" + t + "');");
			}
			function popUpSmall(URL) {
				popUp(URL, 250, 220, 100, 300)
			}
			function popUpMedium(URL) {
				popUp(URL, 560, 300, 100, 200)
			}
			function popUpBig(URL) {
				popUp(URL, 690, 550, 100, 50)
			}
		</script>
	</head>
	<body>
	<!-- #include file="menu\menu.inc" -->
	<form  action="ClaimsViewByServiceID.asp" method="post" name="form">
		<table>
			<tr>
				<th colspan="2">View Claims By Service</th>
			<tr>
			<tr>
				<td>Service</td>
				<td><%= serviceList %></td>
			</tr>
		</table>
		<br/>
		<%= view %>
	</form>
	</body>
</html>