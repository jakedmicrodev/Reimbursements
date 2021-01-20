<!-- #include file="App_Code/ClaimPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim view
Dim startDate
Dim endDate

startDate = Request.Form("StartDate")
endDate = Request.Form("EndDate")

Set manager=New CClaimPageManager

If startDate <> "" And endDate <> "" Then
	view=manager.ViewClaimAmountByDates(startDate, endDate)
End If

Set manager=Nothing
%>
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>View Claims By Insurance Claim Number</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/table.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 		
		<script language="JavaScript">
			function setEndDate()
			{
				var value=document.getElementById("StartDate").value;
				document.getElementById("EndDate").value = value;
			}
		</script>
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<h2>Claim Amounts By Dates</h2>
			<form  action="ClaimAmountsView.asp" method="post" name="form">
				<div class="row">
					<div class="col-10">
						<label for="StartDate">Start Date</label>
					</div>
					<div class="col-15">
						<input type="date" id="StartDate" name="StartDate" size="10" onchange="setEndDate();" />
					</div>
					<div class="col-70">
					</div>
				</div>
				<div class="row">				
					<div class="col-10">
						<label for="EndDate">End Date</label>
					</div>
					<div class="col-15">
						<input type="date" id="EndDate" name="EndDate" size="10" />
					</div>
					<div class="col-70">
					</div>
				</div>
				<div class="row">
					<input type="submit" value="Search">				
				</div>
			</form>
			<div class="row">
			<%= view %>
			</div>
		</div>
	</body>	
</html>