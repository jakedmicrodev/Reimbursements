<!-- #include file="App_Code/StatePageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim stateID

stateID = Request.QueryString("StateID")

Set manager = New CStatePageManager
Set value = manager.SelectStateByID(stateID)

Set manager = Nothing
%>
<html>
	<head>
		<title>Edit State</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form  action="StatePr.asp" method="post" name="form">
				<input type="hidden" name="ProcessType" value="Update">
				<input type="hidden" name="StateID" value="<%= stateID %>">
				<h2>Edit State</h2>
				<div class="row">
					<div class="col-10">
						<label for="StateName">City State</label>
					</div>
					<div class="col-20">
						<input type="text" name="StateName" id="StateName" value="<%= value.StateName %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<input type="submit" value="Save">				
				</div>
			</form>
		</div>
	</body>
</html>
<%
Set value = Nothing
%>