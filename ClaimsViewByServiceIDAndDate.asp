<!-- #include file="App_Code/ClaimPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim view
Dim manager
Dim endDate
Dim startDate
Dim serviceID
Dim serviceList

serviceID = Request.Form("ServiceID")

startDate = Request.Form("StartDate")
If startDate = "" Then startDate = Date()
endDate = Request.Form("EndDate")
If endDate = "" Then endDate = Date()

Set manager=New CClaimPageManager
serviceList = manager.LoadServices(serviceID)

If serviceID <> "" Then
	view=manager.ViewClaimsByServiceIDAndDate(serviceID, startDate, endDate)
End If

'Response.Write("Messages: " & manager.Messages & "<br/>")
Set manager=Nothing
%>
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>View Claims By Service</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/table.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 	
		<script type="text/javascript">
			function submitform() 
			{ 
				document.form.submit(); 
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
			<h2>Claims By Service</h2>
			<form  action="ClaimsViewByServiceIDAndDate.asp" method="post" name="form">
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
						<label for="ServiceID">Service</label>
					</div>
					<div class="col-15">
						<%= serviceList %>
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