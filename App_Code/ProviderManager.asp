<%
Class CProviderManager
	Private mProvider
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
		AddMessage "State: " & mConnection.State, true
		Set mProvider=Nothing
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
	
	Private Function FillProvider(rs)
		Dim value
		Set value=New CProvider

		value.ProviderID=GetFieldValue(rs,"ProviderID","Numeric")
		value.FirstName=GetFieldValue(rs,"FirstName","Text")
		value.Mi=GetFieldValue(rs,"Mi","Text")
		value.LastName=GetFieldValue(rs,"LastName","Text")
		value.Address1=GetFieldValue(rs,"Address1","Text")
		value.Address2=GetFieldValue(rs,"Address2","Text")
		value.CityID=GetFieldValue(rs,"CityID","Numeric")
		value.City=GetFieldValue(rs,"CityName","Text")
		value.StateID=GetFieldValue(rs,"StateID","Numeric")
		value.State=GetFieldValue(rs,"StateName","Text")
		value.Zip=GetFieldValue(rs,"Zip","Text")
		value.Phone=GetFieldValue(rs,"Phone","Text")
			
		Set FillProvider = value
	End Function

	'Public Properties
	Public Property Set Provider(value)
		Set mProvider=value
	End Property
	
	Public Property Get List()
		Set List=mList
	End Property
	
	Public Property Get Messages()
		Messages=mMessages
	End Property
	
	'Public Methods
	Public Sub SelectProviders()
	    'On Error Resume Next
		Dim value
		Dim rs

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelProviders", mConnection
		CheckForError "Select Providers Failed!"

		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillProvider(rs)
			mList.Add CStr(value.ProviderID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
	
	
	Public Sub SelectProvidersByPatientID(patientID)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelProvidersByPatientID " & _
			AddArgument(patientID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Provider by patient ID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillProvider(rs)
			mList.Add CStr(value.ProviderID), value
			rs.MoveNext
		Wend

		rs.Close
		Set rs = Nothing		
	End Sub
		
	Public Sub SelectProviderByID(providerID)
	    On Error Resume Next
		Dim value

		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "pSelProviderByID " & _
			AddArgument(providerID, "NumericTerminal"), mConnection
		CheckForError "Unable to select the Provider by ID!"
		
		Set mList=CreateObject("Scripting.Dictionary")

		While Not rs.EOF
			Set value = FillProvider(rs)
			mList.Add CStr(value.ProviderID), value
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

		sql = "pInsProvider " & _
		AddArgument(mProvider.FirstName, "Text") & _
		AddArgument(mProvider.Mi, "Text") & _
		AddArgument(mProvider.LastName, "Text") & _
		AddArgument(mProvider.Address1, "Text") & _
		AddArgument(mProvider.Address2, "Text") & _
		AddArgument(mProvider.CityID, "Numeric") & _
		AddArgument(mProvider.StateID, "Numeric") & _
		AddArgument(mProvider.Zip, "Text") & _
		AddArgument(mProvider.Phone, "TextTerminal")

		mConnection.Execute sql
		CheckForError "Unable to save new Provider!"
		If mErrorNumber = 0 Then
		    Save=True
		Else
		    Save=False
		End If		
	End Function

	Public Function Update()
		On Error Resume Next
		Dim sql
		
		sql = "pUpdProvider " & _
		AddArgument(mProvider.ProviderID, "Numeric") & _
		AddArgument(mProvider.FirstName, "Text") & _
		AddArgument(mProvider.Mi, "Text") & _
		AddArgument(mProvider.LastName, "Text") & _
		AddArgument(mProvider.Address1, "Text") & _
		AddArgument(mProvider.Address2, "Text") & _
		AddArgument(mProvider.CityID, "Numeric") & _
		AddArgument(mProvider.StateID, "Numeric") & _
		AddArgument(mProvider.Zip, "Text") & _
		AddArgument(mProvider.Phone, "TextTerminal")
		
		mConnection.Execute sql
		CheckForError "Unable to update Provider!"
		If mErrorNumber = 0 Then
		    Update=True
		Else
		    Update=False
		End If
	End Function
End Class
%>