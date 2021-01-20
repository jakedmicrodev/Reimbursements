<!-- #include file="ExpenseTypeManager.asp" -->
<!-- #include file="ExpenseType.asp" -->
<!-- #include file="Common.asp" -->
<%
Class CExpenseTypePageManager
	Private mMessages

 	'Use this for debugging AddMessage "message", true
	Private Sub AddMessage(message, add)
		If add Then
			mMessages=mMessages & message & "<br/>"
		End If
	End Sub
	
	' Private Function IIf(expression, trueValue, falseValue)
		' If expression Then
			' IIf = trueValue
		' Else
			' IIf = falseValue
		' End If
	' End Function
	
	Private Function LoadList(list)
		Dim i
		Dim cls
		Dim keys
		Dim value
		Dim output
		
		keys=list.Keys
		
		output="<table>"
		output=output & "<tr>"
		output=output & "<th>Expense Type</th>"
		output=output & "<th></th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			Set value=list.Item(keys(i))
			output=output & "<tr>"
			output=output & "<td>" & value.ExpenseTypeName & "</td>"
			output=output & "<td>" 
			output=output & "<a href='EditExpenseType.asp?ExpenseTypeID=" & value.ExpenseTypeID & "' class='image' title=''><img alt='Edit' src='images/pencil.svg' width='16' height='16' border='0' /></a>"
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

	Public Function LoadExpenseTypes(expenseTypeID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CExpenseTypeManager		
		manager.SelectExpenseTypes
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='ExpenseTypeID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ExpenseTypeID = CInt(expenseTypeID), " selected ", "")
		    output=output & "<option value='" & value.ExpenseTypeID & "'" & selected & ">" & value.ExpenseTypeName & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadExpenseTypes=output
		
		Set ExpenseType = Nothing
		Set list = Nothing
	End Function
	
	Public Function ViewExpenseTypes()
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CExpenseTypeManager		
		manager.SelectExpenseTypes
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewExpenseTypes = LoadList(list)
		
		Set list=Nothing
	End Function
	
	Public Function ViewExpenseTypeByID(expenseTypeID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CExpenseTypeManager		
		manager.SelectExpenseTypeByID expenseTypeID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewExpenseTypeByID = LoadList(list)
		
		Set list=Nothing
	End Function
		
	Public Function SelectExpenseTypeByID(expenseTypeID)
		Dim manager
		Dim list
		Dim value
		Dim i
		Dim keys
		
		Set manager=New CExpenseTypeManager
		manager.SelectExpenseTypeByID(expenseTypeID)
		
		Set list=manager.List
		Set manager=Nothing
		
		keys=list.Keys
		Set value=list.Item(keys(0))
		Set list=Nothing
		
		Set SelectExpenseTypeByID=value 
	End Function
	
	Public Function Save(value)
		Dim manager
		
		Set manager=New CExpenseTypeManager
		Set manager.ExpenseType=value
		
		Save=manager.Save()
		Set manager=Nothing
	End Function
	
	Public Function Update(value)
		Dim manager
		
		Set manager=New CExpenseTypeManager
		Set manager.ExpenseType=value
		
		Update=manager.Update()
		Set manager=Nothing
	End Function
	
End Class
%>