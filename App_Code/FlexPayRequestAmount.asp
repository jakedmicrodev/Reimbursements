<%
Class CFlexPayRequestAmount
	Private mExpenseType
	Private mStartDate
	Private mEndDate
	Private mExpenseAmount
	Private mInsuranceClaimNumber
	
	'Constructor
	Private Sub Class_Initialize()
		mExpenseType = ""
		mStartDate = ""
		mEndDate = ""
		mExpenseAmount = 0
		mInsuranceClaimNumber = ""
	End Sub
	
	'Public Properties	
	Public Property Let ExpenseType(value)
		mExpenseType=value
	End Property
	Public Property Get ExpenseType()
		ExpenseType=CStr(mExpenseType)
	End Property
	
	Public Property Let StartDate(value)
		mStartDate=value
	End Property
	Public Property Get StartDate()
		StartDate=mStartDate
	End Property

	Public Property Let EndDate(value)
		mEndDate=value
	End Property
	Public Property Get EndDate()
		EndDate=mEndDate
	End Property

	Public Property Let ExpenseAmount(value)
		mExpenseAmount=value
	End Property
	Public Property Get ExpenseAmount()
		ExpenseAmount=CDbl(mExpenseAmount)
	End Property
	
	Public Property Let InsuranceClaimNumber(value)
		mInsuranceClaimNumber=Trim(value)
	End Property
	Public Property Get InsuranceClaimNumber()
		InsuranceClaimNumber=mInsuranceClaimNumber
	End Property
End Class
%>