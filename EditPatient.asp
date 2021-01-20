<!-- #include file="App_Code/PatientPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim value
Dim manager
Dim patientID

patientID = Request.QueryString("PatientID")
Set manager = New CPatientPageManager
Set value = manager.SelectPatientByID(patientID)

Set manager = Nothing
%>
<html>
	<head>
		<title>Edit Patient</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form  action="PatientPr.asp" method="post" name="form">
			<input type="hidden" name="ProcessType" value="Update">
			<input type="hidden" name="PatientID" value="<%= patientID %>">			
				<h2>Edit Patient</h2>
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
					<input type="submit" value="Save">				
				</div>
			</form>
		</div>
	</body>
</html>