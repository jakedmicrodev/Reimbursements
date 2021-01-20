<%
Class CPayee
	Private mPayeeID
	Private mPayeeName
	Private mAddress1
	Private mAddress2
	Private mCityID
	Private mCityName
	Private mStateID
	Private mStateName
	Private mZipCode
	Private mPhoneNumber
	Private mAccountNumber
	Private mNameOnAccount
	Private mActiveCD
	
	'Constructor
	Private Sub Class_Initialize()
		mPayeeID = 0
		mPayeeName = ""
		mAddress1 = ""
		mAddress2 = ""
		mCityID = 0
		mCity = ""
		mStateID = 0
		mState = ""
		mZipCode = ""
		mPhoneNumber = ""
		mAccountNumber = ""
		mNameOnAccount = ""
		mActiveCD = 1
	End Sub
	
	'Public Properties
	Public Property Let PayeeID(value)
		mPayeeID=value
	End Property
	Public Property Get PayeeID()
		PayeeID=mPayeeID
	End Property
	
	Public Property Let PayeeName(value)
		mPayeeName=value
	End Property
	Public Property Get PayeeName()
		PayeeName=mPayeeName
	End Property
	
	Public Property Let Address1(value)
		mAddress1=value
	End Property
	Public Property Get Address1()
		Address1=mAddress1
	End Property
	
	Public Property Let Address2(value)
		mAddress2=value
	End Property
	Public Property Get Address2()
		Address2=mAddress2
	End Property

	Public Property Let CityID(value)
		mCityID=value
	End Property
	Public Property Get CityID()
		CityID=mCityID
	End Property

	Public Property Let CityName(value)
		mCityName=value
	End Property
	Public Property Get CityName()
		CityName=mCityName
	End Property

	Public Property Let StateID(value)
		mStateID=value
	End Property
	Public Property Get StateID()
		StateID=mStateID
	End Property

	Public Property Let StateName(value)
		mStateName=value
	End Property
	Public Property Get StateName()
		StateName=mStateName
	End Property

	Public Property Let ZipCode(value)
		mZipCode=value
	End Property
	Public Property Get ZipCode()
		ZipCode=mZipCode
	End Property

	Public Property Let PhoneNumber(value)
		mPhoneNumber=value
	End Property
	Public Property Get PhoneNumber()
		PhoneNumber=mPhoneNumber
	End Property

	Public Property Let AccountNumber(value)
		mAccountNumber=value
	End Property
	Public Property Get AccountNumber()
		AccountNumber=mAccountNumber
	End Property

	Public Property Let NameOnAccount(value)
		mNameOnAccount=value
	End Property
	Public Property Get NameOnAccount()
		NameOnAccount=mNameOnAccount
	End Property

	Public Property Let ActiveCD(value)
		mActiveCD=value
	End Property
	Public Property Get ActiveCD()
		ActiveCD=mActiveCD
	End Property

	Public Property Get Active()
		If ActiveCD Then
			Active="Yes"
		Else
			Active="No"
		End If
	End Property

	' Public Property Get Active()
		' If ActiveCD = 1 Then
			' Active="Yes"
		' Else
			' Active="No"
		' End If
	' End Property

End Class
%>