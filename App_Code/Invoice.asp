<%
Class CInvoice
	Private mInvoiceID
	Private mInvoiceNumber
	Private mPayeeID
	Private mPayeeName
	Private mAccountID
	Private mAccountNumber
	Private mAmount
	Private mDueDate
	Private mDatePaid
	Private mAmountPaid
	Private mClaimID
	Private mClaimNumber
	
	'Constructor
	Private Sub Class_Initialize()
		mInvoiceID = 0
		mInvoiceNumber = ""
		mPayeeID = 0
		mPayeeName = ""
		mAccountID=0
		mAccountNumber = ""
		mAmount = 0
		mDueDate = ""
		mDatePaid = "1/1/1900"
		mAmountPaid = 0
		mClaimID = 0
		mClaimNumber = ""
	End Sub
	
	'Public Properties
	Public Property Let InvoiceID(value)
		mInvoiceID=value
	End Property
	Public Property Get InvoiceID()
		InvoiceID=mInvoiceID
	End Property
	
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
	
	Public Property Let AccountNumber(value)
		mAccountNumber=value
	End Property
	Public Property Get AccountNumber()
		AccountNumber=mAccountNumber
	End Property
	
	Public Property Let AccountID(value)
		mAccountID=value
	End Property
	Public Property Get AccountID()
		AccountID=mAccountID
	End Property
	
	Public Property Let InvoiceNumber(value)
		mInvoiceNumber=value
	End Property
	Public Property Get InvoiceNumber()
		InvoiceNumber=mInvoiceNumber
	End Property

	Public Property Let Amount(value)
		mAmount=value
	End Property
	Public Property Get Amount()
		Amount=mAmount
	End Property

	Public Property Let DueDate(value)
		mDueDate=value
	End Property
	Public Property Get DueDate()
		DueDate=mDueDate
	End Property

	Public Property Let DatePaid(value)
		mDatePaid=value
	End Property
	Public Property Get DatePaid()
		DatePaid=mDatePaid
	End Property
	
	Public Property Let AmountPaid(value)
		mAmountPaid=value
	End Property
	Public Property Get AmountPaid()
		AmountPaid=mAmountPaid
	End Property			

	Public Property Let ClaimID(value)
		mClaimID=value
	End Property
	Public Property Get ClaimID()
		ClaimID=mClaimID
	End Property
	
	Public Property Let ClaimNumber(value)
		mClaimNumber=value
	End Property
	Public Property Get ClaimNumber()
		ClaimNumber=mClaimNumber
	End Property
End Class
%>