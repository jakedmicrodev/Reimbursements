<!-- #include file="App_Code/ExpenseTypePageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim expenseTypeID

expenseTypeID = Request.QueryString("ExpenseTypeID")

Set manager = New CExpenseTypePageManager
Set value = manager.SelectExpenseTypeByID(expenseTypeID)

Set manager = Nothing
%>
<html>
	<head>
		<title>Edit Expense Type</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form  action="ExpenseTypePr.asp" method="post" name="form"">
				<input type="hidden" name="ExpenseTypeID" value="<%= expenseTypeID %>">
				<input type="hidden" name="ProcessType" value="Update">
				<h2>Edit Expense Type</h2>
				<div class="row">
					<div class="col-15">
						<label for="ExpenseTypeName">Expense Type Name</label>
					</div>
					<div class="col-20">
						<input type="text" name="ExpenseTypeName" id="ExpenseTypeName" value="<%= value.ExpenseTypeName %>">
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