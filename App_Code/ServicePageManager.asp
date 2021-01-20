<!-- #include file="ServiceManager.asp" -->
<!-- #include file="Service.asp" -->
<!-- #include file="Common.asp" -->
<%
Class CServicePageManager
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
		output=output & "<th>Service</th>"
		output=output & "<th></th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			Set value=list.Item(keys(i))
			output=output & "<tr>"
			output=output & "<td>" & value.ServiceName & "</td>"
			output=output & "<td>" 
			output=output & "<a href='EditService.asp?ServiceID=" & value.ServiceID & "' class='image' title=''><img alt='Edit' src='images/pencil.svg' width='16' height='16' border='0' /></a>"
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

	Public Function LoadServices(ServiceID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CServiceManager		
		manager.SelectServices
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='ServiceID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ServiceID = CInt(ServiceID), " selected ", "")
		    output=output & "<option value='" & value.ServiceID & "'" & selected & ">" & value.ServiceName & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadServices=output
		
		Set Service = Nothing
		Set list = Nothing
	End Function
	
	Public Function ViewServices()
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CServiceManager		
		manager.SelectServices
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewServices = LoadList(list)
		
		Set list=Nothing
	End Function
	
	Public Function ViewServiceByID(ServiceID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CServiceManager		
		manager.SelectServiceByID ServiceID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewServiceByID = LoadList(list)
		
		Set list=Nothing
	End Function
		
	Public Function SelectServiceByID(ServiceID)
		Dim manager
		Dim list
		Dim value
		Dim i
		Dim keys
		
		Set manager=New CServiceManager
		manager.SelectServiceByID(ServiceID)
		
		Set list=manager.List
		Set manager=Nothing
		
		keys=list.Keys
		Set value=list.Item(keys(0))
		Set list=Nothing
		
		Set SelectServiceByID=value 
	End Function
	
	Public Function Save(value)
		Dim manager
		
		Set manager=New CServiceManager
		Set manager.Service=value
		
		Save=manager.Save()
		Set manager=Nothing
	End Function
	
	Public Function Update(value)
		Dim manager
		
		Set manager=New CServiceManager
		Set manager.Service=value
		
		Update=manager.Update()
		Set manager=Nothing
	End Function
	
End Class
%>