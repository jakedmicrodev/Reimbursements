<!-- #include file="App_Code/MedicationPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim value
Dim manager
Dim medicationID

medicationID = Request.QueryString("MedicationID")
Set manager = New CMedicationPageManager
Set value = manager.SelectMedicationByID(medicationID)

Set manager = Nothing
%>
<html>
	<head>
		<title>Edit Medication</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form  action="MedicationPr.asp" method="post" name="form">
			<input type="hidden" name="ProcessType" value="Update">
			<input type="hidden" name="MedicationID" value="<%= medicationID %>">			
				<h2>Edit Medication</h2>
				<div class="row">
					<div class="col-10">
						<label for="MedicationName">Medication Name</label>
					</div>
					<div class="col-20">
						<input type="text" name="MedicationName" id="MedicationName" value="<%= value.MedicationName %>">
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