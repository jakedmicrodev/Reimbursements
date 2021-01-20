<%
Class CCategory
	Private mCategoryID
	Private mCategoryName
	
	'Constructor
	Private Sub Class_Initialize()
		mCategoryID = 0
		mCategoryName = ""
	End Sub
	
	'Public Properties
	Public Property Let CategoryID(value)
		mCategoryID=value
	End Property
	Public Property Get CategoryID()
		CategoryID=mCategoryID
	End Property
	
	Public Property Let CategoryName(value)
		mCategoryName=value
	End Property
	Public Property Get CategoryName()
		CategoryName=mCategoryName
	End Property
End Class
%>