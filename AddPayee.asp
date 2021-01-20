<!-- #include file="App_Code/PayeePageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim cityList
Dim stateList

Set manager = New CPayeePageManager
cityList = manager.LoadCities(0)
stateList = manager.LoadStates(0)

Set manager = Nothing
%>
<html>
	<head>
		<title>Add Payee</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form  action="PayeePr.asp" method="post" name="form">
				<input type="hidden" name="ProcessType" value="Add">
				<input type="hidden" name="PayeeID" value="0">
				<h2>Add Payee</h2>
				<div class="row">
					<div class="col-10">
						<label for="PayeeName">Payee Name</label>
					</div>
					<div class="col-20">
						<input type="text" name="PayeeName" id="PayeeName" value="">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="Address1">Address 1</label>
					</div>
					<div class="col-20">
						<input type="text" name="Address1" id="Address1" value="">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="Address2">Address 2</label>
					</div>
					<div class="col-20">
						<input type="text" name="Address2" id="Address2" value="">
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
						<input type="text" name="ZipCode" id="ZipCode" maxlength="9" value="">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="PhoneNumber">Phone Number</label>
					</div>
					<div class="col-20">
						<input type="text" name="PhoneNumber" id="PhoneNumber" maxlength="10" value="">
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
						<input type="text" name="NameOnAccount" id="NameOnAccount" value="">
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