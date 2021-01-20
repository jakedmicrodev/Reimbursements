<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim amount
Dim datePaid
Dim invoiceID

invoiceID = Request.QueryString("InvoiceID")
amount = Request.QueryString("Amount")
datePaid = Request.QueryString("DatePaid")
%>
<html>
	<head>
		<title>Add InvoicePayment</title>
		<link rel="stylesheet" href="css/small_form.css" type="text/css" /> 
	</head>
	<body>
		<div class="container">
			<form  action="InvoicePaymentPr.asp" method="post" name="form">
				<input type="hidden" name="ProcessType" value="Add">
				<input type="hidden" name="InvoiceID" value="<%= invoiceID %>">	
				<h2>Add Invoice Payment</h2>
				<div class="row">
					<div class="col-25">
						<label for="Amount">Amount</label>
					</div>
					<div class="col-50">
						<input type="text" name="Amount" id="Amount" size="5" value="<%= amount %>">
					</div>
				</div>
				<div class="row">
					<div class="col-25">
						<label for="DatePaid">DatePaid</label>
					</div>
					<div class="col-50">
						<input type="date" name="DatePaid" id="DatePaid" value="<%= datePaid %>">
					</div>
				</div>
				<div class="row">
					<input type="submit" value="Save">				
				</div>
			</form>
		</div>
	</body>
<!--			
			<table>
				<tr>
					<th colspan="2">Add Invoice Payment</th>
				<tr>
				<tr>
					<td>Amount</td>
					<td><input name="Amount" size="5" value="<%= amount %>"></td>
				</tr>
				<tr>
					<td>Date Paid</td>
					<td>
						<input type="text" id="DatePaid" name="DatePaid" size="10" value="<%= datePaid %>"> <input type="button" id="trigger" value="..." />
						<script type="text/javascript">
						  Calendar.setup(
							{
							  inputField  : "DatePaid", // ID of the input field
							  ifFormat    : "%m/%d/%Y",    // the date format
							  button      : "trigger"      // ID of the button
							}
						  );
						</script>							
					</td>
				</tr>
				<tr>
					<td class="rowalt" colspan="2" align="left"><input type="submit" value=":: Save ::" /></td>
				</tr>
			</table>
		</form>
	</body>
-->
</html>
<%
Set value = Nothing
%>