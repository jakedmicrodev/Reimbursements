<%
Class CClaim
	Private mClaimID
	Private mProvider
	Private mPatient
	Private mService
	Private mExpenseType
	Private mExpenseDate
	Private mExpenseAmount
	Private mClaimNumber
	Private mInsuranceClaimNumber
	Private mPaidCD
	Private mMedication
	Private mMedicationAmount
	Private mPayee
	Private mInvoiceID
	Private mInvoiceNumber
	
	'Constructor
	Private Sub Class_Initialize()
		mClaimID = 0
		Set mExpenseType = New CExpenseType
		Set mPatient = New CPatient
		Set mProvider = New CProvider
		Set mService = New CService
		Set mMedication = New CMedication
		Set mPayee = New CPayee
		mExpenseDate = ""
		mExpenseAmount = 0
		mClaimNumber = ""
		mInsuranceClaimNumber = ""
		mPaidCD = 0
		mMedicationAmount = 0
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(mExpenseType) Then Set mExpenseType = Nothing
		If IsObject(mPatient) Then Set mPatient = Nothing
		If IsObject(mProvider) Then Set mConnection = Nothing		
		If IsObject(mService) Then Set mService = Nothing
		If IsObject(mMedication) Then Set mMedication = Nothing
		If IsObject(mPayee) Then Set mPayee = Nothing
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
	
	Public Property Let ExpenseTypeID(value)
		mExpenseType.ExpenseTypeID=value
	End Property
	Public Property Get ExpenseTypeID()
		ExpenseTypeID=mExpenseType.ExpenseTypeID
	End Property
	
	Public Property Let ExpenseTypeName(value)
		mExpenseType.ExpenseTypeName=value
	End Property
	Public Property Get ExpenseTypeName()
		ExpenseTypeName=mExpenseType.ExpenseTypeName
	End Property
	
	Public Property Let MedicationID(value)
		mMedication.MedicationID=value
	End Property
	Public Property Get MedicationID()
		MedicationID=mMedication.MedicationID
	End Property
	
	Public Property Let MedicationName(value)
		mMedication.MedicationName=value
	End Property
	Public Property Get MedicationName()
		MedicationName=mMedication.MedicationName
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

	Public Property Let PayeeID(value)
		mPayee.PayeeID=value
	End Property	
	Public Property Get PayeeID()
		PayeeID=mPayee.PayeeID
	End Property
	
	Public Property Let PayeeName(value)
		mPayee.PayeeName=value
	End Property	
	Public Property Get PayeeName()
		PayeeName=mPayee.PayeeName
	End Property
	
	Public Property Let ServiceID(value)
		mService.ServiceID=value
	End Property	
	Public Property Get ServiceID()
		ServiceID=mService.ServiceID
	End Property
	
	Public Property Let ServiceName(value)
		mService.ServiceName=value
	End Property	
	Public Property Get ServiceName()
		ServiceName=mService.ServiceName
	End Property
	
	Public Property Let ExpenseDate(value)
		mExpenseDate=value
	End Property
	Public Property Get ExpenseDate()
		ExpenseDate=mExpenseDate
	End Property

	Public Property Let ExpenseAmount(value)
		mExpenseAmount=value
	End Property
	Public Property Get ExpenseAmount()
		ExpenseAmount=mExpenseAmount
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
	
	Public Property Let MedicationAmount(value)
		mMedicationAmount=value
	End Property
	Public Property Get MedicationAmount()
		MedicationAmount=mMedicationAmount
	End Property	

	Public Property Let InvoiceID(value)
		mInvoiceID=value
	End Property
	Public Property Get InvoiceID()
		InvoiceID=mInvoiceID
	End Property
	
	Public Property Let InvoiceNumber(value)
		mInvoiceNumber=value
	End Property
	Public Property Get InvoiceNumber()
		InvoiceNumber=mInvoiceNumber
	End Property
	
End Class
%>