<!-- #include file="App_Code/InvoicePageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim payeeID
Dim claimID
Dim claimList
Dim accountList
Dim payeeList

payeeID = Request.Form("PayeeID")
claimID = Request.Form("ClaimID")

Set manager = New CInvoicePageManager
payeeList = manager.LoadPayeesWithOnChange(payeeID)

If payeeID <> "" Then
	accountList = manager.LoadAccountsByPayeeID(payeeID, 0)
End If
If payeeID <> "" Then
	claimList = manager.LoadClaimsWithoutInvoiceByPayeeID(payeeID, 0)
End If
Set manager = Nothing
%>
<html>
	<head>
		<title>Add Invoice</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" type="text/css" media="all" href="jscalendar-1.0/skins/aqua/theme.css" title="Aqua" />
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
			<form  action="AddInvoice.asp" method="post" name="form1">
				<h2>Add Invoice</h2>
				<div class="row">
					<div class="col-10">
						<label for="PayeeID">Payee</label>
					</div>
					<div class="col-20">
						<%= payeeList %>
					</div>
					<div class="col-70"></div>
				</div>
			</form>		
			<form action="InvoicePr.asp" method="post" name="form">
				<input type="hidden" name="InvoiceID" value="0">
				<input type="hidden" name="PayeeID" value="<%= payeeID %>">
				<input type="hidden" name="ProcessType" value="Add">
				<div class="row">
					<div class="col-10">
						<label for="AccountNumberID">Account Number</label>
					</div>
					<div class="col-20">
						<%= accountList %>
					</div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ClaimNumber">Claim Number</label>
					</div>
					<div class="col-20">
						<%= claimList %>
					</div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="InvoiceNumber">Invoice Number</label>
					</div>
					<div class="col-20">
						<input type="text" name="InvoiceNumber" id="InvoiceNumber" value="">
					</div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="Amount">Amount</label>
					</div>
					<div class="col-20">
						<input type="text" name="Amount" id="Amount" value="">
					</div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="DueDate">Due Date</label>
					</div>
					<div class="col-20">
						<input type="date" name="DueDate" id="DueDate" value="">
					</div>
				</div>
				<div class="row">
					<input type="submit" value="Save">				
				</div>
			</form>
		</div>
	</body>
</html>