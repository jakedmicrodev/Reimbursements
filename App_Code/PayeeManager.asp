<%
Class CPayeeManager
	Private mPayee
	Private mList
	Private mMessages
	Private mConnection
	Private mErrorNumber
	
	'Constructor
	Private Sub Class_Initialize()
		On Error Resume Next

		mMessages=""
		mErrorNumber=0
		Set mConnection = Server.CreateObject("ADODB.Connection")
		mConnection.Open ConnectionString("MedicalOld")
		' mConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("db\Reimbursements.mdb") & ";User Id=admin;Password=;"
		'mConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("db\WordDefs.accdb") & ";Persist Security Info=False;"
		CheckForError "Open Connection Failed!"
		AddMessage "State: " & mConnection.State, false
		Set mPayee=Nothing
		Set mList=Nothing
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(mConnection) Then
			If mConnection.State = 1 Then
				mConnection.Close
			End If
			Set mConnection = Nothing
		End If
		
		If IsObject(mList) Then Set mList = Nothing
	End Sub

    Private Sub CheckForError(message)
        If Err.number > 0 Then
			mErrorNumber=Err.number
            SetCustomError message
            Err.Clear
        End If
    End Sub

    ' Sub SetCustomError(strMessage)
        ' Display custom message and information from VBScript Err object.

        ' AddMessage "<br/>" & strMessage & "<br/>" & _
          ' "Number (dec) : " & Err.Number & "<br/>" & _
          ' "Number (hex) : &H" & Hex(Err.Number) & "<br/>" & _
          ' "Description  : " & Err.Description & "<br/>" & _
          ' "Source       : " & Err.Source, true
        ' Err.Clear
    ' End Sub

 	'Use this for debugging AddMessage "message", true
	Private Sub AddMessage(message, add)
		If add Then
			mMessages=mMessages & message & "<br/>"
		End If
	End Sub

    ' Function StripApostrophe(text)
        ' StripApostrophe = replace (text,"'","&acute;")
    ' End Function

	' Private Function AddArgument(value, argType)
		' Dim myValue
		
	    ' Select Case argType
	        ' Case "Numeric"
	            ' myValue = value & ", "
	        ' Case "NumericTerminal"
	            ' Nothing to do just return the number
				' myValue = value
	        ' Case "Text"
	            ' myValue = "'" & StripApostrophe(value) & "', "
	        ' Case "TextTerminal"
	            ' myValue = "'" & StripApostrophe(value) & "'"
	    ' End Select
	    
	    ' AddArgument=myValue
	' End Function
	
	' Private Function GetFieldValue(myRS, field, fieldType)
	    ' Dim value
	    
	    ' Select Case fieldType
	        ' Case "Numeric"
	            ' If IsNull(myRS(field)) Then value=0 Else value=myRS(field) End If
	        ' Case Else '"Text"
	            ' If IsNull(myRS(field)) Then value="" Else value=Trim(myRS(field)) End If
	    ' End Select    
	    
	    ' GetFieldValue=value
	' End Function
	
	Private Function FillPayee(rs)
		Dim value
		Set value=New CPayee

		value.PayeeID=GetFieldValue(rs,"PayeeID","Numeric")
		value.PayeeName=GetFieldValue(rs,"PayeeName","Text")
		value.Address1=GetFieldValue(rs,"Address1","Text")
		value.Address2=GetFieldValue(rs,"Address2","Text")
		value.CityID=GetFieldValue(rs,"CityID","Numeric")
		value.CityName=GetFieldValue(rs,"CityName","Text")
		value.StateID=GetFieldValue(rs,"StateID","Numeric")
		value.StateName=GetFieldValue(rs,"StateName","Text")
		value.ZipCode=GetFieldValue(rs,"ZipCode","Text")
		value.PhoneNumber=GetFieldValue(rs,"PhoneNumber","Text")
		value.NameOnAccount=GetFieldValue(rs,"NameOnAccount","Text")
		value.ActiveCD=GetFieldValue(rs,"ActiveCD","Numeric")
			
		Set FillPayee = value
	End Function

	'Public Properties
	Public Property Set Payee(value)
		Set mPayee=value
	End Property
	
	Public Property Get List()
		Set List=mList
	End Property
	
	Public Property Get Messages()
		Messages=mMessages
	End Property
	
	'Public Methods
	Public Sub SelectPayees()
	    'On Error Resume Next
		Dim value
		Dim rs

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelPayees", mConnection
		CheckForError "Select Payees Failed!"

		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillPayee(rs)
			mList.Add CStr(value.PayeeID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub

	Public Sub SelectPayeeByID(payeeID)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelPayeeByID " & _
			AddArgument(payeeID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Payee by ID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillPayee(rs)
			mList.Add CStr(value.PayeeID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
		
	'Microsoft JET Database Engine (0x80004005) Operation must use an updateable query.
	'Modified the permissions for all users to have full control
	Public Function Save()
		'On Error Resume Next
		Dim sql

		sql = "pInsPayee " & _
		AddArgument(mPayee.PayeeName, "Text") & _
		AddArgument(mPayee.Address1, "Text") & _
		AddArgument(mPayee.Address2, "Text") & _
		AddArgument(mPayee.CityID, "Numeric") & _
		AddArgument(mPayee.StateID, "Numeric") & _
		AddArgument(mPayee.ZipCode, "Text") & _
		AddArgument(mPayee.PhoneNumber, "Text") & _
		AddArgument(mPayee.NameOnAccount, "Text") & _
		AddArgument(mPayee.ActiveCD, "NumericTerminal")

		mConnection.Execute sql
		CheckForError "Unable to save new Payee!"
		If mErrorNumber = 0 Then
		    Save=True
		Else
		    Save=False
		End If		
	End Function

	Public Function Update()
		'On Error Resume Next
		Dim sql
		
		sql = "pUpdPayee " & _
		AddArgument(mPayee.PayeeID, "Numeric") & _
		AddArgument(mPayee.PayeeName, "Text") & _
		AddArgument(mPayee.Address1, "Text") & _
		AddArgument(mPayee.Address2, "Text") & _
		AddArgument(mPayee.CityID, "Numeric") & _
		AddArgument(mPayee.StateID, "Numeric") & _
		AddArgument(mPayee.ZipCode, "Text") & _
		AddArgument(mPayee.PhoneNumber, "Text") & _
		AddArgument(mPayee.NameOnAccount, "Text") & _
		AddArgument(mPayee.ActiveCD, "NumericTerminal")
		
		mConnection.Execute sql
		CheckForError "Unable to update Payee!"
		If mErrorNumber = 0 Then
		    Update=True
		Else
		    Update=False
		End If
	End Function
End Class
%>