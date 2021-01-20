<!-- #include file="MedicationManager.asp" -->
<!-- #include file="Medication.asp" -->
<!-- #include file="Common.asp" -->
<%
Class CMedicationPageManager
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
		output=output & "<th>Medication</th>"
		output=output & "<th></th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			Set value=list.Item(keys(i))
			output=output & "<tr>"
			output=output & "<td>" & value.MedicationName & "</td>"
			output=output & "<td>" 
			output=output & "<a href='EditMedication.asp?MedicationID=" & value.MedicationID & "' class='image' title=''><img alt='Edit' src='images/pencil.svg' width='16' height='16' border='0' /></a>"
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

	Public Function LoadMedications(MedicationID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CMedicationManager		
		manager.SelectMedications
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='MedicationID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.MedicationID = CInt(MedicationID), " selected ", "")
		    output=output & "<option value='" & value.MedicationID & "'" & selected & ">" & value.MedicationName & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadMedications=output
		
		Set Medication = Nothing
		Set list = Nothing
	End Function
	
	Public Function ViewMedications()
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CMedicationManager		
		manager.SelectMedications
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewMedications = LoadList(list)
		
		Set list=Nothing
	End Function
	
	Public Function ViewMedicationByID(MedicationID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CMedicationManager		
		manager.SelectMedicationByID MedicationID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewMedicationByID = LoadList(list)
		
		Set list=Nothing
	End Function
		
	Public Function SelectMedicationByID(MedicationID)
		Dim manager
		Dim list
		Dim value
		Dim i
		Dim keys
		
		Set manager=New CMedicationManager
		manager.SelectMedicationByID(MedicationID)
		
		Set list=manager.List
		Set manager=Nothing
		
		keys=list.Keys
		Set value=list.Item(keys(0))
		Set list=Nothing
		
		Set SelectMedicationByID=value 
	End Function
	
	Public Function Save(value)
		Dim manager
		
		Set manager=New CMedicationManager
		Set manager.Medication=value
		
		Save=manager.Save()
		Set manager=Nothing
	End Function
	
	Public Function Update(value)
		Dim manager
		
		Set manager=New CMedicationManager
		Set manager.Medication=value
		
		Update=manager.Update()
		Set manager=Nothing
	End Function
	
End Class
%>