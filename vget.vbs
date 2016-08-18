'-'  wget-like script written in vbscript to get binary files.

Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2

Set objArgs = WScript.Arguments

numArgs = objArgs.Count

If numArgs <> 2 Then
	Wscript.Echo "USAGE: cscript vget.vbs [URL] [OUTPUTFILE]"
	Wscript.Quit(-1)
End If

strURL = objArgs(0)
strOutputFile = objArgs(1)

Set objArgs = Nothing

Set objHTTP = CreateObject("Microsoft.XMLHTTP")
Set binStream = CreateObject("ADODB.Stream")
objHTTP.Open "GET", strURL, False
objHTTP.Send

With binStream
    .Type = adTypeBinary
    .Open
    .Write objHTTP.responseBody
    .SaveToFile strOutputFile, adSaveCreateOverWrite
End With
