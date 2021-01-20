<!-- #include file="App_Code/AccountPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim payeeList
Dim accountID

Set manager = New CAccountPageManager
payeeList = manager.LoadPayees(0)
	
Set manager = Nothing
%>
<html>
	<head>
		<title>Add Account</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">	
			<form  action="AccountPr.asp" method="post" name="form" onload="document.form.PayeeID.focus()">
				<input type="hidden" name="AccountID" value="0">
				<input type="hidden" name="ProcessType" value="Add">
				<h2>Add Account</h2>
				<div class="row">
					<div class="col-10">
						<label for="PayeeID">Payee</label>
					</div>
					<div class="col-20">
						<%= payeeList %>
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="AccountNumber">Account Number</label>
					</div>
					<div class="col-20">
						<input type="text" name="AccountNumber" id="AccountNumber" value="">
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