<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
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
			<h2>Claim Search</h2>
			<div class="row">
				<div class="col-10">
					<img src="images/dv1199021.jpg">
				</div>
				<div class="col-75">
					<form name="form" action="ClaimsSearchView.asp" method="post">
						<div class="row">
							<div class="col-15">
								<input name="searchtype" type="radio" value="rbAmount" border="0" >Amount</input>
							</div>
							<div class="col-20">
								<input type="text" name="Amount" id="Amount" value="" maxlength="7" onFocus="setRadioIndex(document.form.searchtype, 'rbAmount');">
							</div>
							<div class="col-70"></div>
						</div>
						<div class="row">
							<div class="col-15">
								<input name="searchtype" type="radio" value="rbInsuranceID" border="0" >Insurance ID</input>
							</div>
							<div class="col-20">
								<input type="text" name="InsuranceID" id="InsuranceID" value="" size="9" maxlength="10" onFocus="setRadioIndex(document.form.searchtype, 'rbInsuranceID');" />
							</div>
							<div class="col-70"></div>
						</div>
						<div class="row">
							<div class="col-15">
								<input name="searchtype" type="radio" value="rbClaimID" border="0" > Claim ID</input>
							</div>
							<div class="col-20">
								<input type="text" name="ClaimID" id="ClaimID" value="" size="20" maxlength="25" onFocus="setRadioIndex(document.form.searchtype, 'rbClaimID');" />
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
	</body>
</html>
