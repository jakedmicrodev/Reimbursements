<%
Class CCity
	Private mCityID
	Private mCityName
	
	'Constructor
	Private Sub Class_Initialize()
		mCityID = 0
		mCityName = ""
	End Sub
	
	'Public Properties
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
End Class
%>