' Create an instance of the Outlook application
Set MyOutlook = CreateObject("Outlook.Application")
' Get the current Outlook session
Set MySession = MyOutlook.Session

' Retrieve all rules from Outlook and store them in the MyRules variable
Set MyRules = MySession.DefaultStore.GetRules()

' Initialize an array to store rule names for redirection
Dim Array_names_index
Array_names_index = 0

' Declare an array for rule names
' The size is not determined yet, as we don't know how many names will be in it
Dim Array_names()

' Initialize a variable to iterate through the array of names during the delete cycle
Dim Name_index
Name_index = 0

' Check if rules have forwarding to email addresses
For Each MyRule in MyRules
    For Each myAction in MyRule.Actions
        If myAction.Enabled = True Then
            If myAction.ActionType = 6 Then
                ' Add the rule name to the array of names
                wscript.echo("Rule name: " & MyRule.Name & "; Rule execution order: " & MyRule.ExecutionOrder & "Forward to: " & MyRule.Actions.Forward.Recipients.Item(1))
                REDIM PRESERVE Array_names(Array_names_index + 1)
                Array_names(Array_names_index) = MyRule.Name
                Array_names_index = Array_names_index + 1
            End If
        End If
    Next
Next

' Iterate through the names array and delete the corresponding rules
For Name_index = 0 To Array_names_index - 1
    MyRules.Remove(Array_names(Name_index))
    MyRules.Save
Next
