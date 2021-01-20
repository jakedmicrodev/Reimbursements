<%
Class CState
	Private mStateID
	Private mStateName
	
	'Constructor
	Private Sub Class_Initialize()
		mStateID = 0
		mStateName = ""
	End Sub
	
	'Public Properties
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
End Class
%>