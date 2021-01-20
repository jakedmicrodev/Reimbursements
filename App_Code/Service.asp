<%
Class CService
	Private mServiceID
	Private mServiceName
	
	'Constructor
	Private Sub Class_Initialize()
		mServiceID = 0
		mServiceName = ""
	End Sub
	
	'Public Properties
	Public Property Let ServiceID(value)
		mServiceID=value
	End Property
	Public Property Get ServiceID()
		ServiceID=CInt(mServiceID)
	End Property
	Public Property Get ServiceIDToString()
		ServiceIDToString=CStr(mServiceID)
	End Property
	
	Public Property Let ServiceName(value)
		mServiceName=value
	End Property
	Public Property Get ServiceName()
		ServiceName=mServiceName
	End Property
End Class
%>