<!-- #include file="App_Code/ClaimPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim value
Dim view
Dim manager
Dim endDate
Dim startDate
Dim searchBy

Set manager=New CClaimPageManager

Select Case Request.Form("searchtype")
	Case "rbAmount"
		searchBy = "Amount"
		value = Request.Form("Amount")
		view = manager.ViewClaimsByAmount(value)
	Case "rbInsuranceID"
		searchBy = "Insurance ID"
		value = Request.Form("InsuranceID")
		view = manager.ViewClaimsByInsuranceClaimNumber(value)
	Case "rbClaimID"
		searchBy = "Claim ID"
		value = Request.Form("ClaimID")
		view = manager.ViewClaimsByClaimNumber(value)
	Case "rbServiceDate"
		searchBy = "Service Date"
		startDate = Request.Form("StartServiceDate")
		endDate = Request.Form("EndServiceDate")
		view = manager.ViewClaims(startDate, endDate)
	Case Else
		searchBy = "No Search Parameter Selected!"
End Select

Set manager=Nothing
%>
<html>
	<head>
		<title>Claims Search Results</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/table.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 	
		<script language="JavaScript"> 
			function confirmSubmit()
			{
			var agree=confirm("Are you sure you wish to delete this Cricket ACH Return?");
			if (agree)
				return true ;
			else
				return false ;
			}
			function popUp(url, w, h, t, l) {
				day = new Date();
				id = day.getTime();
				eval("page" + id + " = window.open(url, '" + id + "', 'toolbar=0,scrollbars=1,location=0,statusbar=0,menubar=0,resizable=yes,width=" + w + ",height=" + h + ",left=" + l + ",top=" + t + "');");
			}
			function popUpSmall(URL) {
				popUp(URL, 250, 220, 100, 300)
			}
			function popUpMedium(URL) {
				popUp(URL, 560, 300, 100, 200)
			}
			function popUpBig(URL) {
				popUp(URL, 690, 550, 100, 50)
			}
		</script>
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<h5>Claims Search Results by <font color="red"><%= searchBy %></font></h5>
			<a href="ClaimsSearch.asp"><image src="images/lookup.gif" width="16" height="16"> New Search</a>
			<div class="row">
				<%= view %>
			</div>
		</div>		
	</body>
</html>
