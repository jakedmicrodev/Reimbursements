<!-- #include file="AccountManager.asp" -->
<!-- #include file="Account.asp" -->
<!-- #include file="ProviderManager.asp" -->
<!-- #include file="Provider.asp" -->
<!-- #include file="PayeeManager.asp" -->
<!-- #include file="Payee.asp" -->
<!-- #include file="Common.asp" -->
<%
Class CAccountPageManager
	Private mMessages

 	'Use this for debugging AddMessage "message", true
	Private Sub AddMessage(message, add)
		If add Then
			mMessages=mMessages & message & "<br/>"
		End If
	End Sub
		
	Private Function LoadList(list)
		Dim i
		Dim cls
		Dim keys
		Dim value
		Dim output
		
		keys=list.Keys
		
		output="<table>"
		output=output & "<tr>"
		output=output & "<th>Payee</th>"
		output=output & "<th>Account Number</th>"
		output=output & "<th></th>"
		output=output & "</tr>"
		
		For i=0 To list.Count - 1
			' cls = IIf(cls="rowalt", "row", "rowalt")
			
			Set value=list.Item(keys(i))
			output=output & "<tr>"
			output=output & "<td>" & value.PayeeName & "</td>"
			output=output & "<td>" & value.AccountNumber & "</td>"
			output=output & "<td>"
			output=output & "<a href='EditAccount.asp?AccountID=" & value.AccountID & "' class='image' title=''><img alt='Edit' src='images/pencil.svg' width='16' height='16' border='0' /></a>"
			output=output & "</td>"
			output=output & "</tr>"
		Next

		Set value=Nothing
		
		output=output & "</table>"
		
		LoadList = output
	End Function
	
	Public Property Get Messages()
		Messages=mMessages
	End Property

	Public Function IIf(expression, trueValue, falseValue)
		If expression Then
			IIf = trueValue
		Else
			IIf = falseValue
		End If
	End Function

	Public Function LoadAccountsByProviderID(providerID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim i
		
		Set manager=New CAccountManager		
		manager.SelectAccountsByProviderID providerID
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='AccountID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ProviderID = CInt(providerID), " selected ", "")
		    output=output & "<option value='" & value.AccountID & "'" & selected & ">" & value.AccountNumber & "</option>"
	    Next
    	
	    output=output & "</select>"
		
		LoadAccountsByPayeeID=output
		
		Set value = Nothing
		Set list = Nothing
	End Function
	
	Public Function LoadPayees(payeeID)
		On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CPayeeManager		
		manager.SelectPayees
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='PayeeID' id='PayeeID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.PayeeID = CInt(payeeID), " selected ", "")
		    output=output & "<option value='" & value.PayeeID & "'" & selected & ">" & value.PayeeName & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadPayees=output
		
		Set value = Nothing
		Set list = Nothing
		Set manager = Nothing
	End Function
	
	Public Function LoadProviders(providerID)
		'On Error Resume Next
		Dim selected
		Dim manager
		Dim output
		Dim value
		Dim list
		Dim keys
		Dim cls
		Dim i
		
		Set manager=New CProviderManager		
		manager.SelectProviders
		
		Set list=manager.List
		AddMessage "list.Count=" & list.Count, false
		Set manager=Nothing
		
		keys=list.Keys
		
	    output="<select name='ProviderID'>"
		output=output & "<option value='0'>Select</option>"
	    For i=0 To list.Count -1 
		    Set value=list.Item(keys(i))
			selected = IIf(value.ProviderID = CInt(providerID), " selected ", "")
		    output=output & "<option value='" & value.ProviderID & "'" & selected & ">" & value.Name & "</option>"
	    Next
    	
	    output=output & "</select>"
		Set list=Nothing
		Set value=Nothing
				
		LoadProviders=output
		
		Set value = Nothing
		Set list = Nothing
		Set manager = Nothing
	End Function
	
	Public Function SelectAccountByID(accountID)
		Dim manager
		Dim list
		Dim value
		Dim i
		Dim keys
		
		Set manager=New CAccountManager
		manager.SelectAccountByID accountID
		
		Set list=manager.List
		Set manager=Nothing
		
		keys=list.Keys
		Set value=list.Item(keys(0))
		Set list=Nothing
		
		Set SelectAccountByID=value 
	End Function
	
	Public Function ViewAccounts()
		'On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CAccountManager	

		manager.SelectAccounts
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewAccounts = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function ViewAccountsByPayeeID(payeeID)
		On Error Resume Next
		
		Dim manager
		Dim list
		
		Set manager=New CAccountManager		
		manager.SelectAccountsByPayeeID payeeID
		
		Set list=manager.List
		Set manager=Nothing
		
		ViewAccountsByPayeeID = LoadList(list)
		
		Set list=Nothing
	End Function

	Public Function Save(value)
		Dim manager
		
		Set manager=New CAccountManager
		Set manager.Account=value
		
		If manager.Save() Then
			Save=true
		Else
			Save=false
			AddMessage manager.Messages, true
		End If
		
		Set manager=Nothing
	End Function
	
	Public Function Update(value)
		Dim manager
		
		Set manager=New CAccountManager
		Set manager.Account=value
		
		If manager.Update() Then
			Update=true
		Else
			Update=false
			AddMessage manager.Messages, true
		End If
		
		Set manager=Nothing
	End Function
	
End Class
%>