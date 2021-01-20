<!-- #include file="App_Code/ClaimPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim myMonth
Dim myDay
Dim manager
Dim providerList
Dim patientList
Dim serviceList
Dim medicationList
Dim insuranceClaimNumber
Dim expenseTypeList

Const PRESCRIPTION = 6

patientID = Request.Form("PatientID")
providerID = Request.Form("ProviderID")
serviceID = Request.Form("ServiceID")

Set manager = New CClaimPageManager
patientList = manager.LoadPatientsWithOnChange(patientID)

If patientID <> "" Then
	providerList = manager.LoadProvidersByPatientIDWithOnChange(patientID, providerID)
End If

If providerID <> "" Then
	serviceList = manager.LoadServicesByProviderIDWithOnChange(providerID, serviceID)
End If

If serviceID <> "" Then
	If CInt(serviceID) = PRESCRIPTION Then
		medicationList = manager.LoadMedicationsByProviderID(providerID, 0)
	Else
		medicationList = manager.LoadMedications(0)
	End If
End If

expenseTypeList = manager.LoadExpenseTypes(0)
myMonth = manager.IIf(Len(Month(Date())) = 1, "0" & Month(Date()), Month(Date()))
myDay = manager.IIf(Len(Day(Date())) = 1, "0" & Day(Date()), Day(Date()))
insuranceClaimNumber = myMonth & myDay & Year(Date())

Set manager = Nothing
%>
<html>
	<head>
		<title>Add Claim</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
		<script type="text/javascript">
			function submitform()
			{ 
				document.form1.submit(); 
			}
		</script>
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form action="AddClaim.asp" method="post" name="form1">
				<h2>Add Claim</h2>
				<div class="row">
					<div class="col-10">
						<label for="PatientID">Patient</label>
					</div>
					<div class="col-20">
						<%= patientList %>
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ProviderID">Provider</label>
					</div>
					<div class="col-20">
						<%= providerList %>
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ServiceID">Service</label>
					</div>
					<div class="col-20">
						<%= serviceList %>
					</div>
					<div class="col-70"></div>
				</div>
			</form>			
			<form action="ClaimPr.asp" method="post" name="form">
				<input type="hidden" name="ProcessType" value="Add">
				<input type="hidden" name="ClaimID" value="0">
				<input type="hidden" name="PayeeID" value="0">
				<input type="hidden" name="PatientID" value="<%= patientID %>">
				<input type="hidden" name="ProviderID" value="<%= providerID %>">
				<input type="hidden" name="ServiceID" value="<%= serviceID %>">
				<input type="hidden" name="InsuranceClaimNumber" value="<%= insuranceClaimNumber %>">
				<div class="row">
					<div class="col-10">
						<label for="MedicationID">Medication</label>
					</div>
					<div class="col-20">
						<%= medicationList %>
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="MedicationAmount">Medication Amount</label>
					</div>
					<div class="col-20">
						<input type="text" name="MedicationAmount" id="MedicationAmount" size="5" value="">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ExpenseDate">Expense Date</label>
					</div>
					<div class="col-20">
						<input type="date" name="ExpenseDate" id="ExpenseDate" size="10" value="">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ExpenseAmount">Expense Amount</label>
					</div>
					<div class="col-20">
						<input type="text" name="ExpenseAmount" id="ExpenseAmount" size="5" value="">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ExpenseTypeID">Expense Type (Flex Pay)</label>
					</div>
					<div class="col-20">
						<%= expenseTypeList %>
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ClaimNumber">Claim Number</label>
					</div>
					<div class="col-20">
						<input type="text" name="ClaimNumber" id="ClaimNumber" size="20" value="">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="PaidCD">Paid</label>
					</div>
					<div class="col-20">
						<input type="checkbox" name="PaidCD" id="PaidCD">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<input type="submit" value="Save">				
				</div>
				
				<%
				If saved=1 Then
					Response.Write("<font color=green><b>Claim Information Saved</b></font>")
				End If
				%>
			</form>
		</div>
	</body>
</html>