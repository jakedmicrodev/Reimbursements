<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html>
	<head>
		<title>Add City</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form  action="CityPr.asp" method="post" name="form" OnLoad="document.getElementById('CityName').focus();">
				<input type="hidden" name="CityID" value="0">
				<input type="hidden" name="ProcessType" value="Add">
				<h2>Add City</h2>
				<div class="row">
					<div class="col-10">
						<label for="CityName">City Name</label>
					</div>
					<div class="col-20">
						<input type="text" name="CityName" id="CityName" value="">
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