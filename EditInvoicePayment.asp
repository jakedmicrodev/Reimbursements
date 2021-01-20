<!-- #include file="App_Code/InvoicePaymentPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim paymentID
Dim value

paymentID = Request.QueryString("PaymentID")
Set manager = New CInvoicePaymentPageManager
Set value = manager.SelectInvoicePaymentByID(paymentID)
%>
<html>
	<head>
		<title>Edit Invoice Payment</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<form  action="InvoicePaymentPr.asp" method="post" name="form">
				<input type="hidden" name="ProcessType" value="Update">
				<input type="hidden" name="PaymentID" value="<%= paymentID %>">			
				<input type="hidden" name="InvoiceID" value="<%= value.InvoiceID %>">			
				<h2>Edit Invoice Payment</h2>
				<div class="row">
					<div class="col-10">
						<label for="Amount">Amount</label>
					</div>
					<div class="col-20">
						<input type="text" name="Amount" id="Amount" size="5" value="<%= value.Amount %>">
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="DatePaid">Date Paid</label>
					</div>
					<div class="col-20">
						<input type="date" name="DatePaid" id="DatePaid" value="<%= value.EditDatePaid %>">
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