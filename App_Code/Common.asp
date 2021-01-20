<%
Public Function IIf(expression, trueValue, falseValue)
    If expression Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function
	
Public Function StripApostrophe(text)
    StripApostrophe = Replace(text,"'","&acute;")
End Function

Public Function SetCustomError(strMessage)
    'Display custom message and information from VBScript Err object.

    SetCustomError = "<br/>" & strMessage & "<br/>" & _
      "Number (dec) : " & Err.Number & "<br/>" & _
      "Number (hex) : &H" & Hex(Err.Number) & "<br/>" & _
      "Description  : " & Err.Description & "<br/>" & _
      "Source       : " & Err.Source
End Function

Public Function AddArgument(value, argType)
    Dim myValue
    
    Select Case argType
        Case "Date"
            ' myValue = "#" & value & "#, "
            myValue = "'" & value & "', "
        Case "DateTerminal"
            ' myValue = "#" & value & "#"
            myValue = "'" & value & "'"
        Case "Numeric"
            myValue = value & ", "
        Case "NumericTerminal"
            'Nothing to do just return the number
            myValue = value
        Case "Text"
            myValue = "'" & StripApostrophe(value) & "', "
        Case "TextTerminal"
            myValue = "'" & StripApostrophe(value) & "'"
    End Select
    
    AddArgument=myValue
End Function

'Use this for debugging AddMessage "message", true
Public Function AddMessage(message, add)
    If add Then
        AddMessage = message & "<br/>"
    End If
End Function

Public Function GetFieldValue(myRS, field, fieldType)
    Dim value
    
    Select Case fieldType
        Case "Numeric"
            If IsNull(myRS(field)) Then value=0 Else value=myRS(field) End If
        Case Else '"Text"
            If IsNull(myRS(field)) Then value="" Else value=Trim(myRS(field)) End If
    End Select    
    
    GetFieldValue=value
End Function

Public Function ConnectionString(catalog)
    'ConnectionString = "Provider=SQLOLEDB;Data Source=bengal;Initial Catalog=" & catalog & ";User ID=vlrequests;Password=Z3^ezDev!!!"
    ConnectionString = "Provider=SQLOLEDB;Data Source=localhost\sqlexpress;Initial Catalog=" & catalog & ";User ID=vlrequests;Password=Z3^ezDev!!!"
End Function

Public Function FormatDate(value)
	Dim yyyy
	Dim mm
	Dim dd
	
	yyyy = CStr(Year(value))
	mm = CStr(Month(value))
	dd = CStr(Day(value))

	If Len(mm) = 1 Then
		mm = "0" & mm
	End If

	If Len(dd) = 1 Then
		dd = "0" & dd
	End If

	FormatDate = yyyy & "-" + mm & "-" + dd
End Function

Public Function FormatBit(value)
	If value Then
		FormatBit = "checked"
	Else
		FormatBit = ""
	End If
End Function

' Public Function CreateJSONArray(list, jsonType)
    ' Dim output
    ' Dim value
    ' Dim keys
    ' Dim i
    
    ' keys = list.Keys
    
    ' Start the JSON array
    ' output="["
    
    ' For i=0 To list.Count - 1			
        ' Set value=list.Item(keys(i))
        
        ' If jsonType = "Generic" Then
            ' output=output & FillGenericJSON(value)
        ' ElseIf jsonType = "Project" Then
            ' output=output & FillProjectJSON(value)
        ' ElseIf jsonType = "Task" Then
            ' output=output & FillTaskJSON(value)
        ' Else
            ' output=output & FillGenericJSON(value)
        ' End If
        
        ' output=output & ","
    ' Next
    
    ' Set value = Nothing
    
    ' If Right(output, 1) = "," Then 'If there is a trailing comma
        ' output=Left(output, Len(output) - 1) 'Remove the trailing comma
    ' End If
    ' End the JSON array
    ' output=output & "]"

    ' CreateJSONArray = output
' End Function
    
' Public Function FieldExists(ByVal rs, ByVal fieldName)

    ' On Error Resume Next
    ' FieldExists = rs.Fields(fieldName).Name <> ""
    ' If Err <> 0 Then FieldExists = False
    ' Err.Clear
' End Function

' Public Function FillGeneric(rs, idField, nameField)
    ' On Error Resume Next

    ' Dim value
    ' Set value=New CGeneric
    ' value.Key=GetFieldValue(rs,idField,"Text")
    ' value.Value=GetFieldValue(rs,nameField,"Text")
    
    ' Set FillGeneric = value
' End Function

' Public Function FillGenericJSON(value)
    ' Dim output
    ' output="{"
    ' output=output & """" & "Key" &"""" & ":" & """" & value.Key & """" & ","
    ' output=output & """" & "Value" &"""" & ":" & """" & value.Value & """" & "}"
   
    ' FillGenericJSON = output
' End Function

' Public Function FillProject(rs)
    ' Dim value
    ' Set value=New CProject
    ' value.ProjectID=GetFieldValue(rs,"ProjectID","Numeric")
    ' value.Title=GetFieldValue(rs,"Title","Text")
    ' value.Description=GetFieldValue(rs,"Description","Text")
    ' value.Notes=GetFieldValue(rs,"Notes","Text")
    ' value.PriorityID=GetFieldValue(rs,"PriorityID","Numeric")
    ' value.Priority=GetFieldValue(rs,"Priority","Text")
    ' value.StatusID=GetFieldValue(rs,"StatusID","Numeric")
    ' value.Status=GetFieldValue(rs,"Status","Text")
    ' value.StartDate=GetFieldValue(rs,"StartDate","Text")
    ' value.EndDate=GetFieldValue(rs,"EndDate","Text")
    ' value.EnteredBy=GetFieldValue(rs,"EnteredBy","Text")
    ' value.EnteredDate=GetFieldValue(rs,"EnteredDate","Text")
    ' value.ModifiedBy=GetFieldValue(rs,"ModifiedBy","Text")
    ' value.ModifiedDate=GetFieldValue(rs,"ModifiedDate","Text")
    
    ' Set FillProject = value
' End Function

' Private Function FillProjectJSON(value)
    ' Dim output
    ' output="{"
    ' output=output & """" & "ProjectID" &"""" & ":" & """" & value.ProjectID & """" & ","
    ' output=output & """" & "Title" &"""" & ":" & """" & value.Title & """" & ","
    ' output=output & """" & "Description" &"""" & ":" & """" & value.Description & """" & ","
    ' output=output & """" & "Notes" &"""" & ":" & """" & value.Notes & """" & ","
    ' output=output & """" & "PriorityID" &"""" & ":" & """" & value.PriorityID & """" & ","
    ' output=output & """" & "Priority" &"""" & ":" & """" & value.Priority & """" & ","
    ' output=output & """" & "StatusID" &"""" & ":" & """" & value.StatusID & """" & ","
    ' output=output & """" & "Status" &"""" & ":" & """" & value.Status & """" & ","
    ' output=output & """" & "StartDate" &"""" & ":" & """" & value.StartDate & """" & ","
    ' output=output & """" & "EndDate" &"""" & ":" & """" & value.EndDate & """" & ","
    ' output=output & """" & "EnteredBy" &"""" & ":" & """" & value.EnteredBy & """" & ","
    ' output=output & """" & "EnteredDate" &"""" & ":" & """" & value.EnteredDate & """" & ","
    ' output=output & """" & "ModifiedBy" &"""" & ":" & """" & value.ModifiedBy & """" & ","
    ' output=output & """" & "ModifiedDate" &"""" & ":" & """" & value.ModifiedDate & """" & "}"
   
    ' FillProjectJSON = output
' End Function

' Public Function FillTask(rs)
    ' On Error Resume Next
    ' Dim value
    ' Set value=New CTask
    ' value.TaskID=GetFieldValue(rs,"TaskID","Numeric")
    ' value.ProjectID=GetFieldValue(rs,"ProjectID","Numeric")
    ' value.PriorityID=GetFieldValue(rs,"PriorityID","Numeric")
    ' value.Priority=GetFieldValue(rs,"Priority","Text")
    ' value.StatusID=GetFieldValue(rs,"StatusID","Numeric")
    ' value.Status=GetFieldValue(rs,"Status","Text")
    ' value.Description=GetFieldValue(rs,"Description","Text")
    ' value.Notes=GetFieldValue(rs,"Notes","Text")
    ' value.EnteredBy=GetFieldValue(rs,"EnteredBy","Text")
    ' value.EnteredDate=GetFieldValue(rs,"EnteredDate","Text")
    ' value.ModifiedBy=GetFieldValue(rs,"ModifiedBy","Text")
    ' value.ModifiedDate=GetFieldValue(rs,"ModifiedDate","Text")
    
    ' Set FillTask = value
' End Function

' Private Function FillTaskJSON(value)
    ' Dim output
    ' output="{"
    ' output=output & """" & "TaskID" &"""" & ":" & """" & value.TaskID & """" & ","
    ' output=output & """" & "ProjectID" &"""" & ":" & """" & value.ProjectID & """" & ","
    ' output=output & """" & "PriorityID" &"""" & ":" & """" & value.PriorityID & """" & ","
    ' output=output & """" & "Priority" &"""" & ":" & """" & value.Priority & """" & ","
    ' output=output & """" & "StatusID" &"""" & ":" & """" & value.StatusID & """" & ","
    ' output=output & """" & "Status" &"""" & ":" & """" & value.Status & """" & ","
    ' output=output & """" & "Description" &"""" & ":" & """" & value.Description & """" & ","
    ' output=output & """" & "Notes" &"""" & ":" & """" & value.Notes & """" & ","
    ' output=output & """" & "EnteredBy" &"""" & ":" & """" & value.EnteredBy & """" & ","
    ' output=output & """" & "EnteredDate" &"""" & ":" & """" & value.EnteredDate & """" & ","
    ' output=output & """" & "ModifiedBy" &"""" & ":" & """" & value.ModifiedBy & """" & ","
    ' output=output & """" & "ModifiedDate" &"""" & ":" & """" & value.ModifiedDate & """" & "}"
   
    ' FillTaskJSON = output
' End Function

' Public Function GetUserID(value)
    ' Dim parts
    
    ' If InStr(value, "\") Then
        ' parts = Split(value, "\")
        ' GetUserID = parts(1)
    ' Else
        ' GetUserID = value
    ' End If
' End Function


%>