<%
Class CInvoicePayment
	Private mPaymentID
	Private mInvoiceID
	Private mAmount
	Private mDatePaid
	Private mPayeeName
	
	'Constructor
	Private Sub Class_Initialize()
		mPaymentID = 0
		mInvoiceID = 0
		mAmount = 0
		mDatePaid = ""
		mPayeeName = ""
	End Sub
	
	'Public Properties
	Public Property Let PaymentID(value)
		mPaymentID=value
	End Property
	Public Property Get PaymentID()
		PaymentID=mPaymentID
	End Property
	
	Public Property Let InvoiceID(value)
		mInvoiceID=value
	End Property
	Public Property Get InvoiceID()
		InvoiceID=mInvoiceID
	End Property
	
	Public Property Let Amount(value)
		mAmount=value
	End Property
	Public Property Get Amount()
		Amount=mAmount
	End Property
	
	Public Property Let DatePaid(value)
		mDatePaid=value
	End Property
	Public Property Get DatePaid()
		DatePaid=mDatePaid
	End Property
	Public Property Get EditDatePaid()
		EditDatePaid=YEAR(mDatePaid) & _ 
        "-" & Right("0" & Month(mDatePaid),2) & _ 
        "-" & Right("0" & Day(mDatePaid),2) 
	End Property
	
	Public Property Let PayeeName(value)
		mPayeeName=value
	End Property
	Public Property Get PayeeName()
		PayeeName=mPayeeName
	End Property
End Class
%>