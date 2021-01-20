<!-- #include file="ProviderManager.asp" -->
<!-- #include file="Provider.asp" -->
<!-- #include file="CityManager.asp" -->
<!-- #include file="City.asp" -->
<!-- #include file="StateManager.asp" -->
<!-- #include file="State.asp" -->
<!-- #include file="Common.asp" -->
<%
Class CProviderPageManager
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
		output=output & "<th>First Name</th>"
		output=output & "<th>Mi</th>"
		output=output & "<th>Last Name</th>"
		output=output & "<th>Address1</th>"
		output=output & "<th>Address2</th>"
		output=output & "<th>City</th>"
		output=output & "<th>State</th>"
		output=output & "<th>Zip</th>"
		output=output & "<th>Phone</th>"
		output=output & "<th>&nbsp;</th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			Set value=list.Item(keys(i))
			output=output & "<tr>"
			output=output & "<td>" & value.FirstName & "</td>"
			output=output & "<td>" & value.Mi & "</td>"
			output=output & "<td>" & value.LastName & "</td>"
			output=output & "<td>" & value.Address1 & "</td>"
			output=output & "<td>" & value.Address2 & "</td>"
			output=output & "<td>" & value.City & "</td>"
			output=output & "<td>" & value.State & "</td>"
			output=output & "<td>" & value.ZipToString & "</td>"
			output=output & "<td>" & value.PhoneToString & "</td>"
			output=output & "<td>" 
			output=output & "<a href='EditProvider.asp?ProviderID=" & value.ProviderID & "' class='image' title=''><img alt='Edit' src='images/pencil.svg' width='16' height='16' border='0' /></a>"
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
		
		Set manager = Nothing
		Set value = Nothing
		Set list = Nothing
	End Function
	
	Public Function LoadProviders(providerID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CProviderManager		
		manager.SelectProviders
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='ProviderID' id='ProviderID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ProviderID = CInt(providerID), " selected ", "")
		    output=output & "<option value='" & value.ProviderID & "'" & selected & ">" & value.FirstName & IIf(value.Mi <> "", value.Mi & " ", " ") & value.LastName & "</option>"
	    Next
    	
	    output=output & "</select>"
				
		LoadProviders=output
		
		Set manager = Nothing
		Set list=Nothing
		Set value=Nothing
	End Function
	
	Public Function ViewProviders()
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CProviderManager		
		manager.SelectProviders
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewProviders = LoadList(list)
		
		Set list=Nothing
	End Function
	
	Public Function ViewProviderByID(ProviderID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CProviderManager		
		manager.SelectProviderByID ProviderID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewProviderByID = LoadList(list)
		
		Set list=Nothing
	End Function
		
	Public Function SelectProviderByID(ProviderID)
		Dim manager
		Dim list
		Dim value
		Dim i
		Dim keys
		
		Set manager=New CProviderManager
		manager.SelectProviderByID(ProviderID)
		
		Set list=manager.List
		Set manager=Nothing
		
		keys=list.Keys
		Set value=list.Item(keys(0))
		Set list=Nothing
		
		Set SelectProviderByID=value 
	End Function
	
	Public Function Save(value)
		Dim manager
		
		Set manager=New CProviderManager
		Set manager.Provider=value
		
		If Not manager.Save() Then
			AddMessage manager.Messages, true
			Save = False
		Else
			Save = True
		End If
		Set manager=Nothing
	End Function
	
	Public Function Update(value)
		Dim manager
		
		Set manager=New CProviderManager
		Set manager.Provider=value
		
		Update=manager.Update()
		Set manager=Nothing
	End Function
	
End Class
%>