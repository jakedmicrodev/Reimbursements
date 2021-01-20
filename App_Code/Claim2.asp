<%
Class CClaim
	Private mPaidCD
	Private mPatient
	Private mClaimID
	Private mProvider
	Private mClaimNumber
	Private mAccountNumber
	Private mInsuranceClaimNumber
	
	'Constructor
	Private Sub Class_Initialize()
		mPaidCD = 0
		Set mPatient = New CPatient
		mClaimID = 0
		Set mProvider = New CProvider
		mClaimNumber = ""
		mAccountNumber = ""
		mInsuranceClaimNumber = ""
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(mPatient) Then Set mPatient = Nothing
		If IsObject(mProvider) Then Set mConnection = Nothing
    End Sub

	'Public Properties
	Public Property Let ClaimID(value)
		mClaimID=value
	End Property
	Public Property Get ClaimID()
		ClaimID=CInt(mClaimID)
	End Property
	Public Property Get ClaimIDToString()
		ClaimIDToString=CStr(mClaimID)
	End Property
		
	Public Property Let PatientID(value)
		mPatient.PatientID=value
	End Property	
	Public Property Get PatientID()
		PatientID=mPatient.PatientID
	End Property
	
	Public Property Let PatientFirstName(value)
		mPatient.FirstName=value
	End Property	
	Public Property Get PatientFirstName()
		PatientFirstName=mPatient.FirstName
	End Property
		
	Public Property Let PatientMI(value)
		mPatient.MI=value
	End Property	
	Public Property Get PatientMI()
		PatientMI=mPatient.MI
	End Property
		
	Public Property Let PatientLastName(value)
		mPatient.LastName=value
	End Property	
	Public Property Get PatientLastName()
		PatientLastName=mPatient.LastName
	End Property
		
	Public Property Get PatientName()
		PatientName=mPatient.Name
	End Property
		
	Public Property Let ProviderID(value)
		mProvider.ProviderID=value
	End Property	
	Public Property Get ProviderID()
		ProviderID=mProvider.ProviderID
	End Property
	
	Public Property Let ProviderFirstName(value)
		mProvider.FirstName=value
	End Property	
	Public Property Get ProviderFirstName()
		ProviderFirstName=mProvider.FirstName
	End Property
		
	Public Property Let ProviderMI(value)
		mProvider.MI=value
	End Property	
	Public Property Get ProviderMI()
		ProviderMI=mProvider.MI
	End Property
		
	Public Property Let ProviderLastName(value)
		mProvider.LastName=value
	End Property	
	Public Property Get ProviderLastName()
		ProviderLastName=mProvider.LastName
	End Property
		
	Public Property Get ProviderName()
		ProviderName=mProvider.Name
	End Property

	Public Property Let ClaimNumber(value)
		mClaimNumber=value
	End Property
	Public Property Get ClaimNumber()
		ClaimNumber=mClaimNumber
	End Property

	Public Property Let InsuranceClaimNumber(value)
		mInsuranceClaimNumber=value
	End Property
	Public Property Get InsuranceClaimNumber()
		InsuranceClaimNumber=mInsuranceClaimNumber
	End Property

	Public Property Get Paid()
		If PaidCD = 0 Then
			Paid="No"
		Else
			Paid="Yes"
		End If
	End Property
	
	Public Property Let PaidCD(value)
		mPaidCD=value
	End Property
	Public Property Get PaidCD()
		PaidCD=mPaidCD
	End Property
	
	Public Property Get IsPaid()
		If PaidCD = 0 Then
			IsPaid=""
		Else
			IsPaid="checked"
		End If
	End Property
End Class
%>