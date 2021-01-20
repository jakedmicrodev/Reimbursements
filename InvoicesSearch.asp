<!-- #include file="App_Code/InvoicePageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim payeeList
Dim invoiceNumberList

Set manager=New CInvoicePageManager
payeeList = manager.LoadPayeesWithOnFocus(0)
invoiceNumberList = manager.LoadInvoiceNumbersWithOnFocus("")

Set manager = Nothing
%>
<html>
	<head>
		<title>Claims Search</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" type="text/css" media="all" href="jscalendar-1.0/skins/aqua/theme.css" title="Aqua" />
		<link rel="stylesheet" href="css/table.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 		
		<script type="text/javascript" src="jscalendar-1.0/calendar.js"></script>
		<script type="text/javascript" src="jscalendar-1.0/lang/calendar-en.js"></script>
		<script type="text/javascript" src="jscalendar-1.0/calendar-setup.js"></script>
		<script language="JavaScript">
			function setEndDate()
			{
				var value=document.getElementById("StartServiceDate").value;
				document.getElementById("EndServiceDate").value = value;
			}
			
			function validateForm()
			{
				var retVal=true;
				var message="";
				
				if(!validateDate("restoreDate")){
					message += "Restore Date must be mm/dd/yyyy";
					retVal = false;
				}			
				
				if(!retVal)
					alert(message);
 
				return retVal;
			}
			
			function validateDate(textboxID)
			{
				var value=document.getElementById(textboxID).value;
				
				if(value != "")
					return isDate(value);
				else
					return true;
			}
			
			function isDate(value) 
			{
				try {
					//Change the below values to determine which format of date you wish to check. It is set to dd/mm/yyyy by default.
					var MonthIndex = 0;
					var DayIndex = 1;
					var YearIndex = 2;
			 
					value = value.replace(/-/g, "/").replace(/\./g, "/"); 
					var SplitValue = value.split("/");
					var OK = true;
					if (!(SplitValue[DayIndex].length == 1 || SplitValue[DayIndex].length == 2)) {
						OK = false;
					alert("OK1 is " + OK);
					}
					if (OK && !(SplitValue[MonthIndex].length == 1 || SplitValue[MonthIndex].length == 2)) {
						OK = false;
					alert("OK2 is " + OK);
					}
					if (OK && SplitValue[YearIndex].length != 4) {
						OK = false;
					alert("OK3 is " + OK);
					}
					if (OK) {
						var Day = parseInt(SplitValue[DayIndex], 10);
						var Month = parseInt(SplitValue[MonthIndex], 10);
						var Year = parseInt(SplitValue[YearIndex], 10);
						
						if (OK = ((Year > 1900) && (Year <= new Date().getFullYear()))) {
							if (OK = (Month <= 12 && Month > 0)) {
								var LeapYear = (((Year % 4) == 0) && ((Year % 100) != 0) || ((Year % 400) == 0));
			 
								if (Month == 2) {
									OK = LeapYear ? Day <= 29 : Day <= 28;
								}
								else {
									if ((Month == 4) || (Month == 6) || (Month == 9) || (Month == 11)) {
										OK = (Day > 0 && Day <= 30);
									}
									else {
										OK = (Day > 0 && Day <= 31);
									}
								}
							}
						}
					}
					return OK;
				}
				catch (e) {
					return false;
				}
			}
			
			function setRadioIndex ( oRadioGroup, strValue ) {
				var btnArray = oRadioGroup;
				var	btnIndex=0;
				for (var i=0; i<btnArray.length; i++) {
					if (btnArray[i].value==strValue) {
						btnArray[i].checked=true;
						break;
					}
				}		
				return;							 
			}
		</script>
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<h2>Invoice Search</h2>
			<div class="row">
				<div class="col-10">
					<img src="images/dv1199021.jpg">
				</div>
				<div class="col-75">
					<form name="form" action="InvoicesSearchView.asp" method="post">
						<div class="row">
							<div class="col-15">
								<input name="searchtype" type="radio" value="rbAccountNumber" border="0" >Account Number</input>
							</div>
							<div class="col-20">
								<input type="text" name="AccountNumber" value="" size="15" maxlength="20" onFocus="setRadioIndex(document.form.searchtype, 'rbAccountNumber');" />
							</div>
							<div class="col-70"></div>
						</div>
						<div class="row">
							<div class="col-15">
								<input name="searchtype" type="radio" value="rbClaimNumber" border="0" >Claim Number</input>
							</div>
							<div class="col-20">
								<input type="text" name="ClaimNumber" value="" size="15" maxlength="20" onFocus="setRadioIndex(document.form.searchtype, 'rbClaimNumber');" />
							</div>
							<div class="col-70"></div>
						</div>
						<div class="row">
							<div class="col-15">
								<input name="searchtype" type="radio" value="rbPayeeID" border="0" >Payee</input>
							</div>
							<div class="col-20">
								<%= payeeList %>
							</div>
							<div class="col-70"></div>
						</div>
						<div class="row">
							<div class="col-15">
								<input name="searchtype" type="radio" value="rbInvoiceNumber" border="0" >Invoice Number</input>
							</div>
							<div class="col-20">
								<%= invoiceNumberList %>
							</div>
							<div class="col-70"></div>
						</div>
						<div class="row">
							<div class="col-15">
								<input name="searchtype" type="radio" value="rbServiceDate" border="0" >Service Date</input>
							</div>
							<div class="col-20">
								<input name="StartServiceDate" id="StartServiceDate" type="date" size="8" onFocus="setRadioIndex(document.form.searchtype, 'rbServiceDate');" onchange="setEndDate();" />
								<input name="EndServiceDate" id="EndServiceDate" type="date" size="8" onFocus="setRadioIndex(document.form.searchtype, 'rbServiceDate');" />
							</div>
							<div class="col-70"></div>
						</div>
						<div class="row">
							<input type="submit" value="Search">				
						</div>
					</form>
				</div>
			</div>
		</div>
<!--	
		<table width="800" border="0" cellpadding="2" cellspacing="1" class="dotted">
			<tr>
				<td>
					<table width="630" border="0" cellpadding="0" cellspacing="1" align="center" class="withborder">
						<tr>
							<th width="112" rowspan="2" align="center" class="altrow"><img src="images/dv1199021.jpg"></th>
							<th width="518" align="center">Invoices Search</th>
						</tr>
						<tr>
							<td class="altrow" align="center">
								<form name="form" action="InvoicesSearchView.asp" method="post">
									<table width="100%" border="0">
										<tr>
											<td width="37%" align="left"><input name="searchtype" type="radio" value="rbAccountNumber" border="0" > Account Number</td>
											<td width="63%" align="left"><input type="text" name="AccountNumber" value="" size="15" maxlength="20" onFocus="setRadioIndex(document.form.searchtype, 'rbAccountNumber');" /></td>
										</tr>
										<tr>
											<td width="37%" align="left"><input name="searchtype" type="radio" value="rbClaimNumber" border="0" > Claim Number</td>
											<td width="63%" align="left"><input type="text" name="ClaimNumber" value="" size="15" maxlength="20" onFocus="setRadioIndex(document.form.searchtype, 'rbClaimNumber');" /></td>
										</tr>
										<tr>
											<td width="37%" align="left"><input name="searchtype" type="radio" value="rbPayeeID" border="0" > Payee</td>
											<td width="63%" align="left"><%= payeeList %></td>
										</tr>
										<tr>
											<td width="37%" align="left"><input name="searchtype" type="radio" value="rbInvoiceNumber" border="0" > Invoice Number</td>
											<td width="63%" align="left"><%= invoiceNumberList %></td>
										</tr>
										<tr>
											<td width="37%" align="left"><input name="searchtype" type="radio" value="rbServiceDate" border="0" > Service Date:</td>
											<td width="63%" align="left"><input id="StartServiceDate" name="StartServiceDate" type="text" size="8" onFocus="setRadioIndex(document.form.searchtype, 'rbServiceDate');" onchange="setEndDate();" /> <input type="button" id="trigger" value="..." onFocus="setRadioIndex(document.form.searchtype, 'rbServiceDate');">
												<script type="text/javascript">
												Calendar.setup(
													{
												  inputField  : "StartServiceDate", // ID of the input field
												  ifFormat    : "%m/%d/%Y",         // the date format
												  button      : "trigger"           // ID of the button
													}
												  );
												</script> And 
												<input id="EndServiceDate" name="EndServiceDate" type="text" size="8" onFocus="setRadioIndex(document.form.searchtype, 'rbServiceDate');" /> <input type="button" id="trigger1" value="..." onFocus="setRadioIndex(document.form.searchtype, 'rbServiceDate');">
												<script type="text/javascript">
												Calendar.setup(
													{
												  inputField  : "EndServiceDate", // ID of the input field
												  ifFormat    : "%m/%d/%Y",       // the date format
												  button      : "trigger1"        // ID of the button
													}
												  );
												</script> 
											</td>
										</tr>
										<!--
										<tr>
											<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Fixed:</td>
											<td><input type="checkbox" name="FixedCD" id="fixedCD" /></td>
										</tr>
										-->
										<!--
									</table>
									<p>
										<input type="submit" value=":: Go ::">
									</p>
								</form>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		-->
	</body>
</html>
