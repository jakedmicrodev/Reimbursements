<%
Class CAccountManager
	Private mAccount
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
		'mConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("db\Reimbursements.mdb") & ";User Id=admin;Password=;"
		'mConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("db\WordDefs.accdb") & ";Persist Security Info=False;"
		CheckForError "Open Connection Failed!"
		AddMessage "State: " & mConnection.State, true
		Set mAccount=Nothing
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
		Else
			AddMessage "No error found!"
        End If
    End Sub

    Sub SetCustomError(strMessage)
        'Display custom message and information from VBScript Err object.

        AddMessage "<br/>" & strMessage & "<br/>" & _
          "Number (dec) : " & Err.Number & "<br/>" & _
          "Number (hex) : &H" & Hex(Err.Number) & "<br/>" & _
          "Description  : " & Err.Description & "<br/>" & _
          "Source       : " & Err.Source, true
        Err.Clear
    End Sub

 	'Use this for debugging AddMessage "message", true
	Private Sub AddMessage(message, add)
		If add Then
			mMessages=mMessages & message & "<br/>"
		End If
	End Sub

    Function StripApostrophe(text)
        StripApostrophe = replace (text,"'","&acute;")
    End Function

	Private Function AddArgument(value, argType)
		Dim myValue
		
	    Select Case argType
	        Case "Numeric"
	            myValue = value & ", "
	        Case "NumericTerminal"
	            'Nothing to do just return the number
				myValue = value
	        Case "Text"
	            myValue = "'" & StripApostrophe(value) & "', "
	        Case "TextTerminal"
	            myValue = "'" & StripApostrophe(value) & "'"
	        Case "Date"
	            myValue = "#" & value & "#, "
	        Case "DateTerminal"
	            myValue = "#" & value & "#"
	    End Select
	    
	    AddArgument=myValue
	End Function
	
	Private Function GetFieldValue(myRS, field, fieldType)
	    Dim value
	    
	    Select Case fieldType
	        Case "Numeric"
	            If IsNull(myRS(field)) Then value=0 Else value=myRS(field) End If
	        Case Else '"Text"
	            If IsNull(myRS(field)) Then value="" Else value=Trim(myRS(field)) End If
	    End Select    
	    
	    GetFieldValue=value
	End Function
	
	Private Function FillAccount(rs)
		Dim value
		Set value=New CAccount

		value.AccountID=GetFieldValue(rs,"AccountID","Numeric")
		value.AccountNumber=GetFieldValue(rs,"AccountNumber","Text")
		value.PayeeID=GetFieldValue(rs,"PayeeID","Numeric")
		value.PayeeName=GetFieldValue(rs,"PayeeName","Text")
			
		Set FillAccount = value
	End Function

	'Public Properties
	Public Property Set Account(value)
		Set mAccount=value
	End Property
	
	Public Property Get List()
		Set List=mList
	End Property
	
	Public Property Get Messages()
		Messages=mMessages
	End Property
	
	'Public Methods
	Public Sub SelectAccounts()
	    On Error Resume Next
		Dim value
		Dim rs

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelAccounts", mConnection
		CheckForError "Select Accounts Failed!"

		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillAccount(rs)
			mList.Add CStr(value.AccountID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub

	Public Sub SelectAccountByID(accountID)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelAccountByID " & _
			AddArgument(accountID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Account by ID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillAccount(rs)
			mList.Add CStr(value.AccountID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
		
	Public Sub SelectAccountsByPayeeID(payeeID)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelAccountsByPayeeID " & _
			AddArgument(payeeID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Account by PayeeID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillAccount(rs)
			mList.Add CStr(value.AccountID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
		
	Public Sub SelectAccountsByProviderID(providerID)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelAccountsByProviderID " & _
			AddArgument(providerID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Account by ProviderID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillAccount(rs)
			mList.Add CStr(value.AccountID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
		
	'Microsoft JET Database Engine (0x80004005) Operation must use an updateable query.
	'Modified the permissions for all users to have full control
	Public Function Save()
		On Error Resume Next
		Dim sql

		sql = "pInsAccount " & _
		AddArgument(mAccount.PayeeID, "Numeric") & _
		AddArgument(mAccount.AccountNumber, "TextTerminal")

		mConnection.Execute sql
		CheckForError "Unable to save new Account!"
		AddMessage sql, true
		If mErrorNumber = 0 Then
		    Save=true
		Else
		    Save=false
		End If		
	End Function

	Public Function Update()
		On Error Resume Next
		Dim sql
		
		sql = "pUpdAccount " & _
		AddArgument(mAccount.AccountID, "Numeric") & _
		AddArgument(mAccount.PayeeID, "Numeric") & _
		AddArgument(mAccount.AccountNumber, "TextTerminal")
		
		mConnection.Execute sql
		CheckForError "Unable to update Account!"
		If mErrorNumber = 0 Then
		    Update=true
		Else
		    Update=false
		End If
	End Function
End Class
%>