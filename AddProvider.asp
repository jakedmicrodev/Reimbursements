<!-- #include file="App_Code/ProviderPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim message
Dim citiesList
Dim statesList

message = Request.QueryString("error")
Set manager = New CProviderPageManager
citiesList = manager.LoadCities(0)
statesList = manager.LoadStates(0)

Set manager = Nothing
%>
<html>
	<head>
		<title>Add Provider</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form  action="ProviderPr.asp" method="post" name="form">
				<input type="hidden" name="ProcessType" value="Add">
				<input type="hidden" name="ProviderID" value="0">
				<h2>Add Provider</h2>
				<div class="row">
					<div class="col-10">
						<label for="FirstName">First Name</label>
					</div>
					<div class="col-20">
						<input type="text" name="FirstName" id="FirstName" value="">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="Mi">Middle Initial</label>
					</div>
					<div class="col-20">
						<input type="text" name="Mi" id="Mi" maxlength="1" value="">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="LastName">Last Name</label>
					</div>
					<div class="col-20">
						<input type="text" name="LastName" id="LastName" value="">
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
						<%= citiesList %>
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="StateID">State</label>
					</div>
					<div class="col-20">
						<%= statesList %>
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="Zip">Zip Code</label>
					</div>
					<div class="col-20">
						<input type="text" name="Zip" id="Zip" maxlength="9" value="">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="Phone">Phone Number</label>
					</div>
					<div class="col-20">
						<input type="text" name="Phone" id="Phone" maxlength="10" value="">
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