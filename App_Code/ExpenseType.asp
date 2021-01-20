<%
Class CExpenseType
	Private mExpenseTypeID
	Private mExpenseTypeName
	
	'Constructor
	Private Sub Class_Initialize()
		mExpenseTypeID = 0
		mExpenseTypeName = ""
	End Sub
	
	'Public Properties
	Public Property Let ExpenseTypeID(value)
		mExpenseTypeID=value
	End Property
	Public Property Get ExpenseTypeID()
		ExpenseTypeID=CInt(mExpenseTypeID)
	End Property
	Public Property Get ExpenseTypeIDToString()
		ExpenseTypeIDToString=CStr(mExpenseTypeID)
	End Property
	
	Public Property Let ExpenseTypeName(value)
		mExpenseTypeName=value
	End Property
	Public Property Get ExpenseTypeName()
		ExpenseTypeName=CStr(mExpenseTypeName)
	End Property
End Class
%>