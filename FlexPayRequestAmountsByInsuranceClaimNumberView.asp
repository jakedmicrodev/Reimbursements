<!-- #include file="App_Code/ClaimPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim view
Dim claimNumber
Dim claimNumberList

claimNumber = Request.Form("InsuranceClaimNumber")

Set manager=New CClaimPageManager
claimNumberList = manager.LoadInsuranceClaimNumbersWithOnChange(claimNumber)

If claimNumber <> "" Then
	view=manager.ViewFlexPayRequestAmountsByInsuranceClaimNumber(claimNumber)
End If

Set manager=Nothing
%>
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>View Flex Pay Request Amounts</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/table.css" type="text/css" />
		<link rel="stylesheet" href="css/form2.css" type="text/css" /> 		
		<script type="text/javascript">
			function submitform() 
			{ 
				document.form.submit(); 
			}
		</script>
	</head>
	<!-- #include file="menu\menu.inc" -->
	<body>
		<div class="container">
			<h2>Flex Pay Request Amount</h2>
			<form  action="FlexPayRequestAmountsByInsuranceClaimNumberView.asp" method="post" name="form">
				<div class="row">
					<div class="col-20">
						<label for="ClaimNumber">Insurance Claim Number</input>
					</div>
					<div class="col-20">
						<%= claimNumberList %>
					</div>
					<div class="col-70"></div>
				</div>
				<div class="row">
				<%= view %>
				</div>					
			</form>
			<div class="row">
				<div class="col-100">
					<iframe name="ClaimsIDView" width="100%" height="500" src="ClaimsIDView.asp" frameborder="0"></iframe>
				</div>
			</div>
		</div>
	</body>	
<!--		
		<table>
			<tr>
				<th colspan="2">View Flex Pay Request Amounts</th>
			</tr>
			<tr>
				<td>Insurance Claim Number</td>
				<td><%= claimNumberList %></td>
			</tr>
		</table>
		<br/>
		<%= view %>
	</form>
	<br/>
		<iframe name="ClaimsIDView" width="1000" height="300" src="ClaimsIDView.asp" frameborder="0"></iframe>
	</body>
	-->

</html>