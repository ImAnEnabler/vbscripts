Set objArgs = WScript.Arguments

numArgs = objArgs.Count

If numArgs <> 4 Then
	if numArgs <> 6 Then
		Wscript.Echo "Incorrect number of arguments."
		Wscript.Quit(-1)
	End If
End If

For i = 0 To numArgs - 1
	Select Case objArgs(i)
		Case "-f","/f"
			i = i + 1
			strFolder = objArgs(i)
			'Wscript.Echo strFolder
		Case "-e","/e"
			i = i + 1
			strExtension = objArgs(i)
			'Wscript.Echo intThreshold
		Case "-t","/t"
			i = i + 1
			intThreshold = CInt(objArgs(i))
			'Wscript.Echo intThreshold
		Case Else
			'Something not right!

	End Select
Next 

Set objArgs = Nothing

