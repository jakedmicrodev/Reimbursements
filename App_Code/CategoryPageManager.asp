<!-- #include file="CategoryManager.asp" -->
<!-- #include file="Category.asp" -->
<%
Class CCategoryPageManager
	Private mMessages

 	'Use this for debugging AddMessage "message", true
	Private Sub AddMessage(message, add)
		If add Then
			mMessages=mMessages & message & "<br/>"
		End If
	End Sub
	
	Private Function IIf(expression, trueValue, falseValue)
		If expression Then
			IIf = trueValue
		Else
			IIf = falseValue
		End If
	End Function
	
	Private Function LoadList(list)
		Dim i
		Dim cls
		Dim keys
		Dim value
		Dim output
		
		keys=list.Keys
		
		output="<table class='withborder'>"
		output=output & "<tr cellpadding='1' cellspacing='1'>"
		output=output & "<th>Category</th>"
		output=output & "<th></th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			cls = IIf(cls="rowalt", "alt", "rowalt")
			
			Set value=list.Item(keys(i))
			output=output & "<tr class='" & cls & "'>"
			output=output & "<td>" & value.CategoryName & "</td>"
			output=output & "<td>" 
			output=output & "<a href='EditCategory.asp?CategoryID=" & value.CategoryID & "' class='image' title=''><img alt='Edit' src='images/edit.bmp' width='16' height='16' border='0' /></a>"
			output=output & "</td>"
			output=output & "</tr>"
		Next

		Set value=Nothing
		
		output=output & "</table>"
		
		LoadList = output
	End Function
	
	Public Property Get Messages()
		Messages=mMessages
	End Property

	Public Function LoadCategories(categoryID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CCategoryManager		
		manager.SelectCategories
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='CategoryID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.CategoryID = CInt(categoryID), " selected ", "")
		    output=output & "<option value='" & value.CategoryID & "'" & selected & ">" & value.CategoryName & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadCategories=output
		
		Set value = Nothing
		Set list = Nothing
	End Function
	
	Public Function ViewCategories()
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CCategoryManager		
		manager.SelectCategories
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewCategories = LoadList(list)
		
		Set list=Nothing
	End Function
	
	Public Function ViewCategoryByID(categoryID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CCategoryManager		
		manager.SelectCategoryByID categoryID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewCategoryByID = LoadList(list)
		
		Set list=Nothing
	End Function
		
	Public Function SelectCategoryByID(categoryID)
		Dim manager
		Dim list
		Dim value
		Dim i
		Dim keys
		
		Set manager=New CCategoryManager
		manager.SelectCategoryByID(categoryID)
		
		Set list=manager.List
		Set manager=Nothing
		
		keys=list.Keys
		Set value=list.Item(keys(0))
		Set list=Nothing
		
		Set SelectCategoryByID=value 
	End Function
	
	Public Function Save(value)
		Dim manager
		
		Set manager=New CCategoryManager
		Set manager.Category=value
		
		Save=manager.Save()
		Set manager=Nothing
	End Function
	
	Public Function Update(value)
		Dim manager
		
		Set manager=New CCategoryManager
		Set manager.Category=value
		
		Update=manager.Update()
		Set manager=Nothing
	End Function
	
End Class
%>