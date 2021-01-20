<!-- #include file="App_Code/ClaimPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim value
Dim output
Dim manager

value = Request.QueryString("InsuranceClaimNumber")
Set manager = New CClaimPageManager
If value <> "" Then
	output = manager.ViewClaimsByInsuranceClaimNumber(value)
End If

Set manager=Nothing
%>
<html>
	<head>
		<title>Claims By Insurance Claim Number</title>
		<link rel="stylesheet" href="css/table.css" type="text/css" />
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
	<body>
		<h5>Claims By Insurance Claim Number <%= value %></h5>
		<%= output %>
	</body>
</html>
