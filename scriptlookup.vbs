Option Explicit

On Error Resume Next

Dim i
Dim objDict
Dim myArray
Dim dictResults
Dim var1

Set objDict     = CreateObject("Scripting.Dictionary")
Set dictResults = CreateObject("Scripting.Dictionary")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get Arguments
For i = 0 to Wscript.Arguments.Count - 1
	myArray = split(Wscript.Arguments(i),"=",-1,1)
	objDict.Add myArray(0),myArray(1)
Next

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'If objDict.Exists("sender-email") Then
    If objDict.Item("sender-email") <> "" Then
		var1= Mid(objDict.Item("sender-email"),InstrRev(objDict.Item("sender-email"),"@"),1)
			If var1 <> "@" Then				
				dictResults.Add "endpoint-user-name", Mid(objDict.Item("sender-email"),InstrRev(objDict.Item("sender-email"),"/")+1) 	  			
			Else
				dictResults.Add "endpoint-user-name", Mid(objDict.Item("sender-email"),InstrRev(objDict.Item("sender-email"),"@")-7,7)
			End If
	ElseIf objDict.Item("endpoint-user-name") <> "" Then
		  dictResults.Add "endpoint-user-name", Mid(objDict.Item("endpoint-user-name"),InstrRev(objDict.Item("endpoint-user-name"),"\")+1) 	  			
	End If
'End If

If dictResults.Count > 0 Then
	Call DisplayResults()
End If

WScript.Quit(0)
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DisplayResults()

Dim myArray
Dim i
Dim strValue

	myArray = dictResults.Keys  		' Get the keys.
	For i = 0 To dictResults.Count - 1	' Iterate the array.
		strValue = dictResults.item(myArray(i))
		'strValue = "" & strValue & ""
		wscript.echo myArray(i) & "=" & strValue
	Next
	
End Sub