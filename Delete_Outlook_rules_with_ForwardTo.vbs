Set MyOutlook = CreateObject ("Outlook.Application")
Set MySession = MyOutLook.Session
'All rules from Outlook are placed in the MyRules variable
Set MyRules = MySession.DefaultStore.GetRules()

'Variable for future rule name array with redirection
Dim Array_names_index
Array_names_index = 0

'Redirect rule name array
'No size yet, 'cause we don’t know 
'Will there be names in it
Dim Array_names()

'Variable to pass through an array of names in a delete cycle
Dim Name_index
Name_index = 0 

'Each rule is checked for, 
'Does it have a forwarding (forward) to the mail address
For Each MyRule in MyRules
	For Each myAction in MyRule.Actions
		If myAction.Enabled = True Then
			'If ActionType = 6, which corresponds to the Action "Forward to Addressee"
			If myAction.ActionType = 6 Then
				'then the rule name is written to an array of names
				wscript.echo("Rule name: " & MyRule.Name & "; Rule execution order: " & MyRule.ExecutionOrder & "Forward to: " & MyRule.Actions.Forward.Recipients.Item(1))
				REDIM PRESERVE Array_names(Array_names_index+1) 
				Array_names(Array_names_index) = MyRule.Name
				Array_names_index=Array_names_index+1
			End If
		End If
	Next
Next

'Go through the name array and delete the corresponding rules
For Name_index = 0 To Array_names_index-1
	MyRules.Remove(Array_names(Name_index))
	MyRules.Save
Next