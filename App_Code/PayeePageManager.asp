<!-- #include file="PayeeManager.asp" -->
<!-- #include file="Payee.asp" -->
<!-- #include file="CategoryManager.asp" -->
<!-- #include file="Category.asp" -->
<!-- #include file="CityManager.asp" -->
<!-- #include file="City.asp" -->
<!-- #include file="StateManager.asp" -->
<!-- #include file="State.asp" -->
<!-- #include file="Common.asp" -->
<%
Class CPayeePageManager
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
		Dim keys
		Dim value
		Dim output
		
		keys=list.Keys
		
		output="<table>"
		output=output & "<tr>"
		output=output & "<th>Payee</th>"
		output=output & "<th>Address1</th>"
		output=output & "<th>Address2</th>"
		output=output & "<th>City</th>"
		output=output & "<th>State</th>"
		output=output & "<th>Zip</th>"
		output=output & "<th>Phone</th>"
		output=output & "<th>Name On Account</th>"
		output=output & "<th>Active</th>"
		output=output & "<th></th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			Set value=list.Item(keys(i))
			output=output & "<tr>"
			output=output & "<td>" & value.PayeeName & "</td>"
			output=output & "<td>" & value.Address1 & "</td>"
			output=output & "<td>" & value.Address2 & "</td>"
			output=output & "<td>" & value.CityName & "</td>"
			output=output & "<td>" & value.StateName & "</td>"
			output=output & "<td>" & value.ZipCode & "</td>"
			output=output & "<td>" & value.PhoneNumber & "</td>"
			output=output & "<td>" & value.NameOnAccount & "</td>"
			output=output & "<td>" & value.Active & "</td>"
			output=output & "<td>" 
			output=output & "<a href='EditPayee.asp?PayeeID=" & value.PayeeID & "' class='image' title=''><img alt='Edit' src='images/pencil.svg' width='16' height='16' border='0' /></a>"
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

	Public Function LoadPayees(payeeID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CPayeeManager		
		manager.SelectPayees
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='PayeeID' id='PayeeID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.PayeeID = CInt(payeeID), " selected ", "")
		    output=output & "<option value='" & value.PayeeID & "'" & selected & ">" & value.PayeeName & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadPayees=output
		
		Set manager = Nothing
		Set value = Nothing
		Set list = Nothing
	End Function
	
	Public Function LoadStates(stateID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CStateManager		
		manager.SelectStates
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
		output="<select name='StateID' id='StateID'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.StateID = CInt(stateID), " selected ", "")
		    output=output & "<option value='" & value.StateID & "'" & selected & ">" & value.StateName & "</option>"
		Next

	    output=output & "</select>"
				
		LoadStates=output
		
		Set list=Nothing
		Set value=Nothing
	End Function
	
	Public Function LoadCategories(categoryID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
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
				
		LoadCategories=output
		
		Set list=Nothing
		Set value=Nothing
	End Function
	
	Public Function LoadCities(cityID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CCityManager		
		manager.SelectCities
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
		output="<select name='CityID' id='CityID'>"
		output=output & "<option value='0'>Select</option>"
		For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.CityID = CInt(cityID), " selected ", "")
		    output=output & "<option value='" & value.CityID & "'" & selected & ">" & value.CityName & "</option>"
		Next

	    output=output & "</select>"
				
		LoadCities=output
		
		Set list=Nothing
		Set value=Nothing
	End Function
	
	Public Function ViewPayees()
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CPayeeManager		
		manager.SelectPayees
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewPayees = LoadList(list)
		
		Set list=Nothing
	End Function
	
'	Public Function ViewPayeeByID(payeeID)
'		On Error Resume Next
'		
'		Dim manager
'		Dim list
'		
'		Set manager=New CPayeeManager		
'		manager.SelectPayeeByID payeeID
'		
'		Set list=manager.List
'		Set manager=Nothing
'		
'		ViewPayeeByID = LoadList(list)
'		
'		Set list=Nothing
'	End Function
		
	Public Function SelectPayeeByID(payeeID)
		Dim manager
		Dim list
		Dim value
		Dim i
		Dim keys
		
		Set manager=New CPayeeManager
		manager.SelectPayeeByID payeeID
		
		Set list=manager.List
		Set manager=Nothing
		
		keys=list.Keys
		Set value=list.Item(keys(0))
		Set list=Nothing
		
		Set SelectPayeeByID=value 
	End Function
	
	Public Function Save(value)
		Dim manager
		
		Set manager=New CPayeeManager
		Set manager.Payee=value
		
		If manager.Save() Then
			Save=True
		Else
			Save=false
			AddMessage manager.Messages, true
		End If
		
		Set manager=Nothing
	End Function
	
	Public Function Update(value)
		Dim manager
		
		Set manager=New CPayeeManager
		Set manager.Payee=value
		
		If manager.Update() Then
			Update=True
		Else
			Update=false
			AddMessage manager.Messages, true
		End If
		
		Set manager=Nothing
	End Function
	
End Class
%>