<!-- #include file="App_Code/CategoryPageManager.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"
        "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<%
Dim manager
Dim categoryID

categoryID = Request.QueryString("CategoryID")

Set manager = New CCategoryPageManager
Set value = manager.SelectCategoryByID(categoryID)

Set manager = Nothing
%>
<html>
	<head>
		<title>Edit Category</title>
		<link rel="stylesheet" href="cssMenus/menu_style.css" type="text/css" />
	</head>
	<body>
	<!-- #include file="menu\menu.inc" -->
		<form  action="CategoryPr.asp" method="post" name="form">
			<input type="hidden" name="ProcessType" value="Update">
			<input type="hidden" name="CategoryID" value="<%= categoryID %>">
			<table>
				<tr>
					<th colspan="2">Edit Category</th>
				<tr>
				<tr>
					<td>Category Name</td>
					<td><input name="CategoryName" size="30" value="<%= value.CategoryName %>"></td>
				</tr>
				<tr>
					<td class="rowalt" colspan="2" align="left"><input type="submit" value=":: Save ::" /></td>
				</tr>
			</table>
		</form>
	</body>
</html>
<%
Set value = Nothing
%>