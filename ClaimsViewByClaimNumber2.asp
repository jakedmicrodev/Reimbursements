<!-- #include file="App_Code/ClaimPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim view
Dim claimNumber
Dim claimNumberList

claimNumber = "152885124088107999" 'Request.Form("ClaimNumber")

Set manager=New CClaimPageManager
claimNumberList = manager.LoadClaimNumbersWithOnChange(claimNumber)

If claimNumber <> "" Then
	view=manager.ViewClaimsByClaimNumber(claimNumber)
End If

'Response.Write("Messages: " & manager.Messages & "<br/>")
Set manager=Nothing
%>
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>View Claims By Claim Number</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<!--<link rel="stylesheet" href="css/table.css" type="text/css" />-->
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
		<!-- Latest compiled and minified CSS -->
		<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css" integrity="sha512-dTfge/zgoMYpP7QbHy4gWMEGsbsdZeCXz7irItjcC3sPUFtf0kuFbDz/ixG7ArTxmDjLXDmezHubeNikyKGVyQ==" crossorigin="anonymous">

		<!-- Optional theme -->
		<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap-theme.min.css" integrity="sha384-aUGj/X2zp5rLCbBxumKTCw2Z50WgIr1vs/PFN4praOTvYXWlVyh2UtNUU0KAUhAX" crossorigin="anonymous">

		<!-- Latest compiled and minified JavaScript -->
		<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/js/bootstrap.min.js" integrity="sha512-K1qjQ+NcF2TYO/eI3M6v8EiNYZfA95pQumfvcVrTHtwQVDG+aHRqLi/ETn2uB+1JqwYqVG3LIvdm9lj6imS/pQ==" crossorigin="anonymous"></script>	

		<link rel="stylesheet" href="http://cdn.datatables.net/1.10.2/css/jquery.dataTables.min.css"></style>
		<script type="text/javascript" src="http://cdn.datatables.net/1.10.2/js/jquery.dataTables.min.js"></script>
		<script>
		$(document).ready(function(){
			$('#my-table').dataTable();
		});
		</script>
		<!-- #include file="menu\menu.inc" -->
	<form  action="ClaimsViewByClaimNumber.asp" method="post" name="form">
		<table>
			<tr>
				<th colspan="2">View Claims By Claim Number</th>
			<tr>
			<tr>
				<td>Claim Number</td>
				<td><%= claimNumberList %></td>
			</tr>
		</table>
		<br/>
		<div class="container">
		<%= view %>
		</div>
	</form>
	</body>
</html>