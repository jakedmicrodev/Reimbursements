<!-- #include file="App_Code/PayeePageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim cityList
Dim stateList
Dim categoryList
Dim payeeID
Dim value
Dim hasBalance
Dim isChecked

payeeID = Request.QueryString("PayeeID")
Set manager = New CPayeePageManager
Set value = manager.SelectPayeeByID(payeeID)

isChecked = FormatBit(value.ActiveCD)
cityList = manager.LoadCities(value.CityID)
stateList = manager.LoadStates(value.StateID)

Set manager = Nothing
%>
<html>
	<head>
		<title>Edit Payee</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form  action="PayeePr.asp" method="post" name="form">
			<input type="hidden" name="ProcessType" value="Update">
			<input type="hidden" name="PayeeID" value="<%= payeeID %>">			
				<h2>Edit Payee</h2>
				<div class="row">
					<div class="col-10">
						<label for="PayeeName">Payee Name</label>
					</div>
					<div class="col-20">
						<input type="text" name="PayeeName" id="PayeeName" value="<%= value.PayeeName %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="Address1">Address 1</label>
					</div>
					<div class="col-20">
						<input type="text" name="Address1" id="Address1" value="<%= value.Address1 %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="Address2">Address 2</label>
					</div>
					<div class="col-20">
						<input type="text" name="Address2" id="Address2" value="<%= value.Address2 %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="CityID">City</label>
					</div>
					<div class="col-20">
						<%= cityList %>
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="StateID">State</label>
					</div>
					<div class="col-20">
						<%= stateList %>
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ZipCode">Zip Code</label>
					</div>
					<div class="col-20">
						<input type="text" name="ZipCode" id="ZipCode" maxlength="9" value="<%= value.ZipCode %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="PhoneNumber">Phone Number</label>
					</div>
					<div class="col-20">
						<input type="text" name="PhoneNumber" id="PhoneNumber" maxlength="10" value="<%= value.PhoneNumber %>">
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
					<div class="col-10">
						<label for="NameOnAccount">Name On Account</label>
					</div>
					<div class="col-20">
						<input type="text" name="NameOnAccount" id="NameOnAccount" value="<%= value.NameOnAccount %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ActiveCD">Active</label>
					</div>
					<div class="col-20">
						<input type="checkbox" name="ActiveCD" id="ActiveCD" value="<%= value.ActiveCD %>" <%= isChecked %> />
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