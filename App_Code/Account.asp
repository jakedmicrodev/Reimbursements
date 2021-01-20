<%
Class CAccount
	Private mAccountID
	Private mAccountNumber
	Private mPayeeID
	Private mPayeeName
	
	'Constructor
	Private Sub Class_Initialize()
		mAccountID = 0
		mAccountNumber = ""
		mPayeeID = 0
		mPayeeName = ""
	End Sub
	
	'Public Properties
	Public Property Let AccountID(value)
		mAccountID=value
	End Property
	Public Property Get AccountID()
		AccountID=mAccountID
	End Property
	Public Property Get AccountIDToString()
		AccountIDToString=CStr(mAccountID)
	End Property
	
	Public Property Let AccountNumber(value)
		mAccountNumber=value
	End Property
	Public Property Get AccountNumber()
		AccountNumber=mAccountNumber
	End Property
	
	Public Property Let PayeeID(value)
		mPayeeID=value
	End Property
	Public Property Get PayeeID()
		PayeeID=CInt(mPayeeID)
	End Property
	Public Property Get PayeeIDToString()
		PayeeIDToString=CStr(mPayeeID)
	End Property
	
	Public Property Let PayeeName(value)
		mPayeeName=value
	End Property
	Public Property Get PayeeName()
		PayeeName=mPayeeName
	End Property
	
End Class
%>