<!-- #include file="App_Code/CityPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim cityID

cityID = Request.QueryString("CityID")

Set manager = New CCityPageManager
Set value = manager.SelectCityByID(cityID)

Set manager = Nothing
%>
<html>
	<head>
		<title>Edit City</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form  action="CityPr.asp" method="post" name="form"">
				<input type="hidden" name="CityID" value="<%= cityID %>">
				<input type="hidden" name="ProcessType" value="Update">
				<h2>Edit City</h2>
				<div class="row">
					<div class="col-10">
						<label for="CityName">City Name</label>
					</div>
					<div class="col-20">
						<input type="text" name="CityName" id="CityName" value="<%= value.CityName %>">
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