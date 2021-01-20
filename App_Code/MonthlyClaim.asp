<%
Class CMonthlyClaim
	Private mMonthID
	Private mMonth
	Private mYear
	Private mTotal
	
	'Constructor
	Private Sub Class_Initialize()
		mMonthID = 0
		mMonth = ""
		mYear = ""
		mTotal = 0
	End Sub
	
	'Public Properties
	Public Property Let MonthID(value)
		mMonthID=Trim(value)
	End Property
	Public Property Get MonthID()
		MonthID=CInt(mMonthID)
	End Property
	Public Property Get MonthIDToString()
		MonthIDToString=CStr(mMonthID)
	End Property
	
	Public Property Get Month()
		Select Case mMonthID
			Case 1
				Month = "January"
			Case 2
				Month = "February"
			Case 3
				Month = "March"
			Case 4
				Month = "April"
			Case 5
				Month = "May"
			Case 6
				Month = "June"
			Case 7
				Month = "July"
			Case 8
				Month = "August"
			Case 9
				Month = "September"
			Case 10
				Month = "October"
			Case 11
				Month = "November"
			Case 12
				Month = "December"
		End Select
	End Property
	
	Public Property Let Year(value)
		mYear=Trim(value)
	End Property
	Public Property Get Year()
		Year=mYear
	End Property
	
	Public Property Let Total(value)
		mTotal=Trim(value)
	End Property
	Public Property Get Total()
		Total=CDbl(mTotal)
	End Property
End Class
%>