<%
Class CMedication
	Private mMedicationID
	Private mMedicationName
	
	'Constructor
	Private Sub Class_Initialize()
		mMedicationID = 0
		mMedicationName = ""
	End Sub
	
	'Public Properties
	Public Property Let MedicationID(value)
		mMedicationID=value
	End Property
	Public Property Get MedicationID()
		MedicationID=CInt(mMedicationID)
	End Property
	Public Property Get MedicationIDToString()
		MedicationIDToString=CStr(mMedicationID)
	End Property
	
	Public Property Let MedicationName(value)
		mMedicationName=value
	End Property
	Public Property Get MedicationName()
		MedicationName=CStr(mMedicationName)
	End Property
End Class
%>