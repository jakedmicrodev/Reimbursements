<%
Class CProvider
	Private mProviderID
	Private mFirstName
	Private mMi
	Private mLastName
	Private mAddress1
	Private mAddress2
	Private mCityID
	Private mCity
	Private mStateID
	Private mState
	Private mZip
	Private mPhone
	
	'Constructor
	Private Sub Class_Initialize()
		mProviderID = 0
		mFirstName = ""
		mMi = ""
		mLastName = ""
		mAddress1 = ""
		mAddress2 = ""
		mCityID = 0
		mCity = ""
		mStateID = 0
		mState = ""
		mZip = ""
		mPhone = ""
	End Sub
	
	'Public Properties
	Public Property Let ProviderID(value)
		mProviderID=value
	End Property
	Public Property Get ProviderID()
		ProviderID=CInt(mProviderID)
	End Property
	Public Property Get ProviderIDToString()
		ProviderIDToString=CStr(mProviderID)
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

	Public Property Let City(value)
		mCity=value
	End Property
	Public Property Get City()
		City=mCity
	End Property

	Public Property Let StateID(value)
		mStateID=value
	End Property
	Public Property Get StateID()
		StateID=mStateID
	End Property

	Public Property Let State(value)
		mState=value
	End Property
	Public Property Get State()
		State=mState
	End Property

	Public Property Let Zip(value)
		mZip=value
	End Property
	Public Property Get Zip()
		Zip=mZip
	End Property
	Public Property Get ZipToString()
		If Len(mZip) = 9 Then
			ZipToString=Left(mZip, 5) & "-" & Right(mZip, 4)
		Else
			ZipToString=mZip
		End If
	End Property

	Public Property Let Phone(value)
		mPhone=value
	End Property
	Public Property Get Phone()
		Phone=mPhone
	End Property
	Public Property Get PhoneToString()
		If Len(mPhone) = 10 Then
			PhoneToString=Left(mPhone, 3) & "-" & Mid(mPhone, 4, 3) & "-" & Right(mPhone, 4)
		Else
			PhoneToString=mPhone
		End If
	End Property
	
End Class
%>