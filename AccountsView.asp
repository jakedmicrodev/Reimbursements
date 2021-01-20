<!-- #include file="App_Code/AccountPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim view

Set manager=New CAccountPageManager
view=manager.ViewAccounts()

'Response.Write("Messages: " & manager.Messages & "<br/>")
Set manager=Nothing
%>
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>View Accounts</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/table.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 	
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<h2>Accounts</h2>
			<div class="row">
				<%= view %>
			</div>
		</div>		
	</body>
</html>