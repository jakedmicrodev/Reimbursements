<!-- #include file="CityManager.asp" -->
<!-- #include file="City.asp" -->
<!-- #include file="Common.asp" -->
<%
Class CCityPageManager
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
		output=output & "<th>City</th>"
		output=output & "<th></th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			Set value=list.Item(keys(i))
			output=output & "<tr>"
			output=output & "<td>" & value.CityName & "</td>"
			output=output & "<td>" 
			output=output & "<a href='EditCity.asp?CityID=" & value.CityID & "' class='image' title=''><img alt='Edit' src='images/pencil.svg' width='16' height='16' border='0' /></a>"
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
		Dim cls
		Dim i
		
		Set manager=New CCityManager		
		manager.SelectCities
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='CityID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.CityID = CInt(cityID), " selected ", "")
		    output=output & "<option value='" & value.CityID & "'" & selected & ">" & value.CityName & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadCities=output
		
		Set value = Nothing
		Set list = Nothing
	End Function
	
	Public Function ViewCities()
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CCityManager		
		manager.SelectCities
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewCities = LoadList(list)
		
		Set list=Nothing
	End Function
	
	Public Function ViewCityByID(CityID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CCityManager		
		manager.SelectCityByID CityID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewCityByID = LoadList(list)
		
		Set list=Nothing
	End Function
		
	Public Function SelectCityByID(CityID)
		Dim manager
		Dim list
		Dim value
		Dim i
		Dim keys
		
		Set manager=New CCityManager
		manager.SelectCityByID(CityID)
		
		Set list=manager.List
		Set manager=Nothing
		
		keys=list.Keys
		Set value=list.Item(keys(0))
		Set list=Nothing
		
		Set SelectCityByID=value 
	End Function
	
	Public Function Save(value)
		Dim manager
		
		Set manager=New CCityManager
		Set manager.City=value
		
		Save=manager.Save()
		Set manager=Nothing
	End Function
	
	Public Function Update(value)
		Dim manager
		
		Set manager=New CCityManager
		Set manager.City=value
		
		Update=manager.Update()
		Set manager=Nothing
	End Function
	
End Class
%>