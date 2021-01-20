<!-- #include file="App_Code/ClaimPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim view
Dim manager
Dim endDate
Dim startDate
Dim claimType
Dim claimTypeList

startDate = Request.Form("StartDate")
'If startDate = "" Then startDate = Date()
endDate = Request.Form("EndDate")
'If endDate = "" Then endDate = Date()
claimType = Request.Form("ClaimType")

Set manager=New CClaimPageManager
claimTypeList = manager.LoadClaimTypes(claimType)

If startDate <> "" And endDate <> "" Then
	view=manager.ViewClaimsByMonth(startDate, endDate, claimType)
End If

'Response.Write("Messages: " & manager.Messages & "<br/>")
Set manager=Nothing
%>
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>View Claims</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" type="text/css" media="all" href="jscalendar-1.0/skins/aqua/theme.css" title="Aqua" />
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
			<form  action="ClaimsByMonth.asp" method="post" name="form">
				<div class="row">
					<div class="col-10">
						<label for="StartDate">Start Date</label>
					</div>
					<div class="col-15">
						<input type="date" id="StartDate" name="StartDate" size="10" value="<%= startDate %>" onchange="setEndDate();" />
					</div>
					<div class="col-70">
					</div>
				</div>
				<div class="row">				
					<div class="col-10">
						<label for="EndDate">End Date</label>
					</div>
					<div class="col-15">
						<input type="date" id="EndDate" name="EndDate" size="10" value="<%= endDate %>" />
					</div>
					<div class="col-70">
					</div>
				</div>
				<div class="row">
					<div class="col-10">
						<label for="ClaimTypeID">Claim Type</label>
					</div>
					<div class="col-15">
						<%= claimTypeList %>
					</div>
					<div class="col-75"></div>
				</div>
				<div class="row">
					<div class="col-80"></div>
				</div>
				<div class="row">
					<input type="submit" value="Search">				
				</div>
				<div class="row">
				<%= view %>
				</div>
			</form>
		</div>
	</body>
</html>