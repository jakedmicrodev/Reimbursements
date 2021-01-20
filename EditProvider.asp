<!-- #include file="App_Code/ProviderPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim value
Dim manager
Dim providerID
Dim citiesList
Dim statesList

providerID = Request.QueryString("ProviderID")
Set manager = New CProviderPageManager
Set value = manager.SelectProviderByID(providerID)
citiesList = manager.LoadCities(value.CityID)
statesList = manager.LoadStates(value.StateID)

Set manager = Nothing
%>
<html>
	<head>
		<title>Edit Provider</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form  action="ProviderPr.asp" method="post" name="form">
			<input type="hidden" name="ProcessType" value="Update">
			<input type="hidden" name="ProviderID" value="<%= providerID %>">			
				<h2>Edit Provider</h2>
				<div class="row">
					<div class="col-10">
						<label for="FirstName">First Name</label>
					</div>
					<div class="col-20">
						<input type="text" name="FirstName" id="FirstName" value="<%= value.FirstName %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="Mi">Middle Initial</label>
					</div>
					<div class="col-20">
						<input type="text" name="Mi" id="Mi" maxlength="1" value="<%= value.Mi %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="LastName">Last Name</label>
					</div>
					<div class="col-20">
						<input type="text" name="LastName" id="LastName" value="<%= value.LastName %>">
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
						<input type="text" name="Zip" id="Zip" maxlength="9" value="<%= value.Zip %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="Phone">Phone Number</label>
					</div>
					<div class="col-20">
						<input type="text" name="Phone" id="Phone" maxlength="10" value="<%= value.Phone %>">
					</div>
					<div class="col-70"></div>
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
<!--		<form  action="ProviderPr.asp" method="post" name="form">
			<input type="hidden" name="ProcessType" value="Update">
			<input type="hidden" name="ProviderID" value="<%= providerID %>">			
			<table>
				<tr>
					<th colspan="2">Edit Provider</th>
				<tr>
				<tr>
					<td>First Name</td>
					<td><input name="FirstName" size="30" value="<%= value.FirstName %>"></td>
				</tr>
				<tr>
					<td>Middle Initial</td>
					<td><input name="Mi" size="1" value="<%= value.Mi %>" maxlength="1"></td>
				</tr>
				<tr>
					<td>Last Name</td>
					<td><input name="LastName" size="30" value="<%= value.LastName %>"></td>
				</tr>
				<tr>
					<td>Address</td>
					<td><input name="Address1" size="30" value="<%= value.Address1 %>"></td>
				</tr>
				<tr>
					<td>Address2</td>
					<td><input name="Address2" size="30" value="<%= value.Address2 %>"></td>
				</tr>
				<tr>
					<td>City</td>
					<td><%= citiesList %></td>
				</tr>
				<tr>
					<td>State</td>
					<td><%= statesList %></td>
				</tr>
				<tr>
					<td>Zip</td>
					<td><input name="Zip" size="9" value="<%= value.Zip %>"></td>
				</tr>
				<tr>
					<td>Phone</td>
					<td><input name="Phone" size="10" value="<%= value.Phone %>"></td>
				</tr>
				<tr>
					<td class="rowalt" colspan="2" align="left"><input type="submit" value=":: Save ::" /></td>
				</tr>
			</table>
		</form>
	</body>
	-->
</html>