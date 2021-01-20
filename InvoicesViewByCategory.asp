<!-- #include file="App_Code/InvoicePageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim view
Dim manager
Dim categoryID
Dim categoryList

categoryID = Request.Form("CategoryID")

Set manager=New CInvoicePageManager
categoryList = manager.LoadCategoriesWithOnChange(categoryID)

If categoryID <> "" Then
	view=manager.ViewInvoicesBycategoryID(categoryID)
End If

'Response.Write("Messages: " & manager.Messages & "<br/>")
Set manager=Nothing
%>
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<title>View Invoices</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
		<link rel="stylesheet" href="css/table.css" type="text/css" />
		<script type="text/javascript">
			function submitform() 
			{ 
				document.form.submit(); 
			}
		</script>
	</head>
	<body>
	<!-- #include file="menu\menu.inc" -->
	<form  action="InvoicesViewByCategory.asp" method="post" name="form">
		<table>
			<tr>
				<th colspan="2">View Invoices By Category</th>
			<tr>
			<tr>
				<td>Category</td>
				<td><%= categoryList %></td>
			</tr>
		</table>
		<br/>
		<%= view %>
	</form>
	</body>
</html>