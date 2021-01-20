<!-- #include file="App_Code/ClaimPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim checked
Dim manager
Dim claimID
Dim payeeList
Dim providerList
Dim patientList
Dim serviceList
Dim medicationList
Dim expenseTypeList
Dim expenseDate
Dim isChecked

claimID = Request.QueryString("ClaimID")

Set manager = New CClaimPageManager
Set value = manager.SelectClaimByID(claimID)

payeeList = manager.LoadPayees(value.PayeeID)
providerList = manager.LoadProviders(value.ProviderID)
patientList = manager.LoadPatients(value.PatientID)
serviceList = manager.LoadServices(value.ServiceID)
medicationList = manager.LoadMedications(value.MedicationID)
expenseTypeList = manager.LoadExpenseTypes(value.ExpenseTypeID)

expenseDate = FormatDate(value.ExpenseDate)
isChecked = FormatBit(value.PaidCD)

Set manager = Nothing
%>
<html>
	<head>
		<title>Edit Claim</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form  action="ClaimPr.asp" method="post" name="form">
				<input type="hidden" name="ProcessType" value="Update">
				<input type="hidden" name="ClaimID" value="<%= claimID %>">
				<h2>Edit Claim</h2>
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
						<label for="PayeeID">Payee</label>
					</div>
					<div class="col-20">
						<%= payeeList %>
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
						<input type="text" name="MedicationAmount" id="MedicationAmount" size="5" value="<%= value.MedicationAmount %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ExpenseDate">Expense Date</label>
					</div>
					<div class="col-20">
						<input type="date" name="ExpenseDate" id="ExpenseDate" value="<%= expenseDate %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ExpenseAmount">Expense Amount</label>
					</div>
					<div class="col-20">
						<input type="text" name="ExpenseAmount" id="ExpenseAmount" size="5" value="<%= value.ExpenseAmount %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-15">
						<label for="ExpenseTypeID">Expense Type (Flex Pay)</label>
					</div>
					<div class="col-15">
						<%= expenseTypeList %>
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ClaimNumber">Claim Number</label>
					</div>
					<div class="col-20">
						<input type="text" name="ClaimNumber" id="ClaimNumber" size="20" value="<%= value.ClaimNumber %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-12">
						<label for="InsuranceClaimNumber">Insurance Claim Number</label>
					</div>
					<div class="col-18">
						<input type="text" name="InsuranceClaimNumber" id="InsuranceClaimNumber" size="20" value="<%= value.InsuranceClaimNumber %>">
					</div>
					<div class="col-60"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="PaidCD">Paid</label>
					</div>
					<div class="col-20">
						<input type="checkbox" name="PaidCD" id="PaidCD" value="<%= value.PaidCD %>" <%= isChecked %>>
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