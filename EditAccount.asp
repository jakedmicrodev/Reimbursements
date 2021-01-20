<!-- #include file="App_Code/AccountPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim payeeList
Dim accountID

accountID = Request.QueryString("AccountID")

Set manager = New CAccountPageManager
Set value = manager.SelectAccountByID(accountID)
payeeList = manager.LoadPayees(value.PayeeID)

Set manager = Nothing
%>
<html>
	<head>
		<title>Edit Account</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">	
			<form  action="AccountPr.asp" method="post" name="form"">
				<input type="hidden" name="AccountID" value="<%= accountID %>">
				<input type="hidden" name="ProcessType" value="Update">
				<h2>Edit Account</h2>
				<div class="row">
					<div class="col-10">
						<label for="PayeeID">Payee</label>
					</div>
					<div class="col-25">
						<%= payeeList %>
					</div>
					<div class="col-65"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="AccountNumber">Account Number</label>
					</div>
					<div class="col-15">
						<input type="text" name="AccountNumber" id="AccountNumber" value="<%= value.AccountNumber %>">
					</div>
					<div class="col-75"></div>
				</div>
				<div class="row">
					<input type="submit" value="Save">				
				</div>
			</form>
		</div>
	</body>
	<!--
	</head>
	<body>
	<!-- #include file="menu\menu.inc" -->
		<!--
		<form  action="AccountPr.asp" method="post" name="form">
			<input type="hidden" name="ProcessType" value="Update">
			<input type="hidden" name="AccountID" value="<%= accountID %>">			
			<table>
				<tr>
					<th colspan="2">Edit Account</th>
				<tr>
				<tr>
					<td>Payee</td>
					<td><%= payeeList %></td>
				</tr>
				<tr>
					<td>Account Number</td>
					<td><input name="AccountNumber" size="15" value="<%= value.AccountNumber %>"></td>
				</tr>
				<tr>
					<td class="rowalt" colspan="2" align="left"><input type="submit" value=":: Save ::" /></td>
				</tr>
			</table>
		</form>
	</body>
	 -->
</html>
<%
Set value = Nothing
%>