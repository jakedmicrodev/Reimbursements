<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html>
	<head>
		<title>Add Service</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form  action="ServicePr.asp" method="post" name="form">
				<input type="hidden" name="ProcessType" value="Add">
				<h2>Add Service</h2>
				<div class="row">
					<div class="col-10">
						<label for="ServiceName">Service Name</label>
					</div>
					<div class="col-20">
						<input type="text" name="ServiceName" id="ServiceName" value="">
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