<!-- #include file="StateManager.asp" -->
<!-- #include file="State.asp" -->
<!-- #include file="Common.asp" -->
<%
Class CStatePageManager
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
		output=output & "<th>State</th>"
		output=output & "<th></th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			Set value=list.Item(keys(i))
			output=output & "<tr>"
			output=output & "<td>" & value.StateName & "</td>"
			output=output & "<td>" 
			output=output & "<a href='EditState.asp?StateID=" & value.StateID & "' class='image' title=''><img alt='Edit' src='images/pencil.svg' width='16' height='16' border='0' /></a>"
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

	Public Function LoadStates(stateID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
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
		Set list=Nothing
		Set value=Nothing
				
		LoadStates=output
	End Function
	
	Public Function ViewStates()
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CStateManager
		manager.SelectStates
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewStates = LoadList(list)
		
		Set list=Nothing
	End Function
	
	Public Function ViewStateByID(stateID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CStateManager		
		manager.SelectStateByID stateID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewStateByID = LoadList(list)
		
		Set list=Nothing
	End Function
		
	Public Function SelectStateByID(stateID)
		Dim manager
		Dim list
		Dim value
		Dim i
		Dim keys
		
		Set manager=New CStateManager
		manager.SelectStateByID stateID
		
		Set list=manager.List
		Set manager=Nothing
		
		keys=list.Keys
		Set value=list.Item(keys(0))
		Set list=Nothing
		
		Set SelectStateByID=value 
	End Function
	
	Public Function Save(value)
		Dim manager
		
		Set manager=New CStateManager
		Set manager.State=value
		
		Save=manager.Save()
		Set manager=Nothing
	End Function
	
	Public Function Update(value)
		Dim manager
		
		Set manager=New CStateManager
		Set manager.State=value
		
		Update=manager.Update()
		Set manager=Nothing
	End Function
	
End Class
%>