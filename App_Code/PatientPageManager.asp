<!-- #include file="PatientManager.asp" -->
<!-- #include file="Patient.asp" -->
<!-- #include file="Common.asp" -->
<%
Class CPatientPageManager
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
		'output=output & "<thead>"		
		output=output & "<tr>"
		output=output & "<th>First Name</th>"
		output=output & "<th>Mi</th>"
		output=output & "<th>Last Name</th>"
		output=output & "<th></th>"
		output=output & "</tr>"
		'output=output & "</thead>"		
		
		'output=output & "<tbody>"
		For i=0 To list.Count - 1			
			Set value=list.Item(keys(i))
			output=output & "<tr>"
			output=output & "<td>" & value.FirstName & "</td>"
			output=output & "<td>" & value.Mi & "</td>"
			output=output & "<td>" & value.LastName & "</td>"
			output=output & "<td>" 
			output=output & "<a href='EditPatient.asp?PatientID=" & value.PatientID & "' class='image' title=''><img alt='Edit' src='images/pencil.svg' width='16' height='16' border='0' /></a>"
			output=output & "</td>"
			output=output & "</tr>"
		Next
		'output=output & "</tbody>"		

		Set value=Nothing
		
		output=output & "</table>"
		
		LoadList = output
	End Function
	
	Public Property Get Messages()
		Messages=mMessages
	End Property

	Public Function LoadPatients(patientID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CPatientManager		
		manager.SelectPatients
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='PatientID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.PatientID = CInt(patientID), " selected ", "")
		    output=output & "<option value='" & value.PatientID & "'" & selected & ">" & value.FirstName & IIf(value.Mi <> "", value.Mi & " ", " ") & value.LastName & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadPatients=output
		
		Set Patient = Nothing
		Set list = Nothing
	End Function
	
	Public Function ViewPatients()
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CPatientManager		
		manager.SelectPatients
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewPatients = LoadList(list)
		
		Set list=Nothing
	End Function
	
	Public Function ViewPatientByID(patientID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CPatientManager		
		manager.SelectPatientByID patientID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewPatientByID = LoadList(list)
		
		Set list=Nothing
	End Function
		
	Public Function SelectPatientByID(patientID)
		Dim manager
		Dim list
		Dim value
		Dim i
		Dim keys
		
		Set manager=New CPatientManager
		manager.SelectPatientByID(patientID)
		
		Set list=manager.List
		Set manager=Nothing
		
		keys=list.Keys
		Set value=list.Item(keys(0))
		Set list=Nothing
		
		Set SelectPatientByID=value 
	End Function
	
	Public Function Save(value)
		Dim manager
		
		Set manager=New CPatientManager
		Set manager.Patient=value
		
		Save=manager.Save()
		Set manager=Nothing
	End Function
	
	Public Function Update(value)
		Dim manager
		
		Set manager=New CPatientManager
		Set manager.Patient=value
		
		Update=manager.Update()
		Set manager=Nothing
	End Function
	
End Class
%>