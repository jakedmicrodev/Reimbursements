<!-- #include file="App_Code/InvoicePageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim invoiceID
Dim datePaid
Dim amountPaid
Dim claimList
Dim payeeList
Dim accountList
Dim paymentList
Dim dueDate

invoiceID = Request.QueryString("InvoiceID")

Set manager = New CInvoicePageManager
Set value = manager.SelectInvoiceByID(invoiceID)
payeeList = manager.LoadPayees(value.PayeeID)
accountList = manager.LoadAccountsByPayeeID(value.PayeeID, value.AccountID)
paymentList = manager.ViewInvoicePaymentsByInvoiceID(invoiceID)
claimList = manager.LoadClaimsByPayeeID(value.PayeeID, value.ClaimID)

datePaid = manager.IIf(value.DatePaid <> "1/1/1900", value.DatePaid, "")
amountPaid = manager.IIf(CDbl(value.AmountPaid) <> 0, value.AmountPaid, "")

dueDate = FormatDate(value.DueDate)

Set manager = Nothing
%>
<html>
	<head>
		<title>Edit Invoice</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/table.css" type="text/css" />
		<link rel="stylesheet" type="text/css" media="all" href="jscalendar-1.0/skins/aqua/theme.css" title="Aqua" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
		<script language="JavaScript">
			function confirmSubmit()
			{
			var agree=confirm("Are you sure you wish to delete this payment?");
			if (agree)
				return true ;
			else
				return false ;
			}

			function popUp(URL) {
			day = new Date();
			id = day.getTime();
			eval("page" + id + " = window.open(URL, '" + id + "', 'toolbar=0,scrollbars=0,location=0,statusbar=0,menubar=0,resizable=yes,width=500,height=360,left = 200,top = 100');");
			}
		</script>
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">	
			<form action="InvoicePr.asp" method="post" name="form">
				<input type="hidden" name="ProcessType" value="Update">
				<input type="hidden" name="InvoiceID" value="<%= invoiceID %>">			
				<h2>Edit Invoice</h2>
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
						<label for="AccountID">Account Number</label>
					</div>
					<div class="col-20">
						<%= accountList %>
					</div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ClaimID">Claim Number</label>
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
						<input type="text" name="InvoiceNumber" id="InvoiceNumber" value="<%= value.InvoiceNumber %>">
					</div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="Amount">Amount</label>
					</div>
					<div class="col-20">
						<input type="text" name="Amount" id="Amount" value="<%= value.Amount %>">
					</div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="DueDate">Due Date</label>
					</div>
					<div class="col-20">
						<input type="date" name="DueDate" id="DueDate" value="<%= dueDate %>">
					</div>
				</div>
				<div class="row">
					<div class="col-7">
						<input type="submit" value="Save">
					</div>
					<div class="col-10">
						<input type="button" onclick="javascript:popUp('AddInvoicePayment.asp?InvoiceID=<%= invoiceID %>&Amount=<%= value.Amount %>&DatePaid=<%= Date()%>')" value="Make Payment"/>
					</div>
				</div>
				<div class="row">
					<%= paymentList%>
				</div>
			</form>
		</div>
	</body>
</html>