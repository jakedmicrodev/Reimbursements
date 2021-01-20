<%
Class CPatient
	Private mPatientID
	Private mFirstName
	Private mMi
	Private mLastName
	
	'Constructor
	Private Sub Class_Initialize()
		mPatientID = 0
		mFirstName = ""
		mMi = ""
		mLastName = ""
	End Sub
	
	'Public Properties
	Public Property Let PatientID(value)
		mPatientID=value
	End Property
	Public Property Get PatientID()
		PatientID=CInt(mPatientID)
	End Property
	Public Property Get PatientIDToString()
		PatientIDToString=CStr(mPatientID)
	End Property
	
	Public Property Let FirstName(value)
		mFirstName=value
	End Property
	Public Property Get FirstName()
		FirstName=mFirstName
	End Property
	
	Public Property Let Mi(value)
		mMi=value
	End Property
	Public Property Get Mi()
		Mi=mMi
	End Property
	
	Public Property Let LastName(value)
		mLastName=value
	End Property
	Public Property Get LastName()
		LastName=mLastName
	End Property

	Public Property Get Name()
		Dim value
		
		value = FirstName & " "
		If Mi <> "" Then
			value = value & Mi & " "
		End If
		
		Name=value & LastName
	End Property
End Class
%>