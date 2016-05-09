
'-'
'-'  invokeAllChecks.vbs
'-'
'-'  VBscript implementation of the Invoke-AllChecks function of PowerUp developed by @harmj0y
'-'  by: @ImAnEnabler
'-'  
'-'  In the environment I work, sc.exe is not allowed for non-admins, so I used WMI instead.
'-'  Save the vbs file and run with cscript:
'-'    cscript //nologo invokeAllChecks.vbs
'-'


invokeAllChecks

sub invokeAllChecks()
	Wscript.Echo vbCrLf
	Wscript.Echo "[*] Checking if user is in a local group with administrative privileges..." & vbCrLf
	isAdmin

	Wscript.Echo vbCrLf
	Wscript.Echo "[*] Checking for unquoted service paths..." & vbCrLf
	getServiceUnquoted

	Wscript.Echo vbCrLf
	Wscript.Echo "[*] Checking service executable permissions..." & vbCrLf
	getServiceEXEPerms

	Wscript.Echo vbCrLf
	Wscript.Echo "[*] Checking service permissions..." & vbCrLf
	getServicePerms

	Wscript.Echo vbCrLf
	Wscript.Echo "[*] Checking for unattended install files..." & vbCrLf
	getUnattendedInstallFiles

	Wscript.Echo vbCrLf
	Wscript.Echo "[*] Checking %PATH% for potentially hijackable .dll locations..." & vbCrLf
	invokeFindPathHijack
	
	Wscript.Echo vbCrLf
	Wscript.Echo "[*] Checking for AlwaysInstallElevated registry key..." & vbCrLf
	getRegAlwaysInstallElevated
	
	Wscript.Echo vbCrLf
	Wscript.Echo "[*] Checking for Autologon credentials in registry..." & vbCrLf
	checkAutoAdminLogon
	
	'-' TODO:
	'"[*] Checking for encrypted web.config strings..." & vbCrLf
	'"[*] Checking for encrypted application pool and virtual directory passwords..." & vbCrLf
end sub

sub isAdmin()
	Set objShell = WScript.CreateObject("WScript.Shell")
	'-' Get location of cmd.exe 
	comspec = objShell.ExpandEnvironmentStrings("%comspec%")
	'-' Get groups back from whoami.  I tried many ways to get this through WMI, 
	'-' so that it could be run on XP systems, but was unsuccessful.
	set objResults = objShell.Exec(comspec & " /c whoami.exe /groups")
	Wscript.Sleep 200  '-' it runs async, so lets give it a few milliseconds to run
	strResults = objResults.StdOut.ReadAll
	
	if instr(1, strResults, "S-1-5-32-544", vbtextcompare) > 0 Then ' in local administrators group
		Wscript.Echo "[+] User is in a local group that grants administrative privileges!"
		if instr(1, strResults, "S-1-16-12288", vbtextcompare) > 0 Then ' high-level context = elevated
			Wscript.Echo "[*] You're already running elevated!"
		elseif instr(1, strResults, "S-1-16-8192", vbtextcompare) > 0 Then  ' med-level context = not elevated
			Wscript.Echo "[*] Run a BypassUAC attack to elevate privileges to admin."
		end if
	end if
	set objResults = Nothing
	Set objShell = Nothing
end sub

Sub getServiceUnquoted
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	'-' Get services with unquoted paths
	Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service where NOT PathName LIKE '" & chr(34) & "%'")

	For Each objService in colListOfServices
		'-' check and see if there's a space before the ".exe"
		if (instr(1, objService.PathName, Chr(32), vbTextCompare) > 0) AND _
		 (instr(1, objService.PathName, Chr(32), vbTextCompare) < instr(1, objService.PathName, ".exe", vbTextCompare)) Then
			Wscript.Echo "[+] Unquoted service path: " & objService.Name & " - " & objService.PathName
		end if
	Next
	Set colListOfServices = Nothing
	Set objWMIService = Nothing
end sub

sub getServiceEXEPerms
	Const FILE_WRITE_DATA  = &h000002
	Const FILE_APPEND_DATA = &h000004
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	'-' Get paths to service executables which aren't in system32 folder
	Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service where NOT pathname like '%system32%'")

	For Each objService in colListOfServices
		'-' Get the path through to the ".exe"; if it starts with a quote, drop that off as well
		if (Left(objService.PathName, 1) = """") then 
			objServicePath = mid(objService.PathName, 2, instr(1,objService.PathName, ".exe", vbTextCompare)+2) 
		else 
			objServicePath = mid(objService.PathName, 1, instr(1,objService.PathName, ".exe", vbTextCompare)+3)
		end if
		'-' Get an instance of
		Set objShare = objWMIService.Get("CIM_DataFile.Name='" & objServicePath & "'")
		
		'-' See if the effective permissions say we have write permissions
		isWritable = objShare.GetEffectivePermission(FILE_WRITE_DATA)
		'-' See if the effective permissions say we have append privileges
		isAppendable = objShare.GetEffectivePermission(FILE_APPEND_DATA)
		
		if isWritable then 
			wscript.echo "[+] Vulnerable service executable: " & objServicePath
		end if
		'-' If the file is in use, the write check may fail; if we can append to it, we may still be in luck
		if NOT isWritable AND isAppendable then 
			wscript.echo "[+] Possible vulnerable service executable: " & objServicePath
			wscript.echo objService.State
		end if
	next
	Set objShare = Nothing
	Set colListOfServices = Nothing
	Set objWMIService = Nothing
end sub

sub getServicePerms
	'-' Possible ErrorControl Values to try and set to
	Set dErrCtl = CreateObject("Scripting.Dictionary") 
	dErrCtl.Add "Ignore", 0
	dErrCtl.Add "Normal", 1
	dErrCtl.Add "Severe", 2
	dErrCtl.Add "Critical", 3
	dErrCtl.Add "Unknown", 4
	
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	'-' Get list of services
	Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service")

	For Each objService in colListOfServices
		'-' Try to set the ErrorControl value to the same as it currently is; a return value of 0 is SUCCESS
		If objService.Change( , , , dErrCtl(objService.ErrorControl)) = 0 Then
			Wscript.Echo "[+] Vulnerable service: " & objService.Name & " - " & objService.PathName
		End If
	next
	Set objShare = Nothing
	Set colListOfServices = Nothing
	Set objWMIService = Nothing
	Set dErrCtl = Nothing
end sub

sub getUnattendedInstallFiles
	Set objShell = CreateObject("WScript.Shell")
	windir = objShell.ExpandEnvironmentStrings("%windir%")
	set objShell = Nothing
	'-' List of file locations to check
	arrFiles = array("c:\sysprep\sysprep.xml", _
                    "c:\sysprep\sysprep.inf", _
                    "c:\sysprep.inf", _
                    windir & "\Panther\Unattended.xml", _
                    windir & "\Panther\Unattend\Unattended.xml", _
                    windir & "\Panther\Unattend.xml", _
                    windir & "\Panther\Unattend\Unattend.xml", _
                    windir & "\System32\Sysprep\unattend.xml", _
                    windir & "\System32\Sysprep\Panther\unattend.xml")

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	for i = 0 to ubound(arrFiles)
		if objFSO.FileExists(arrFiles(i)) then
			wscript.echo "[+] Unattended install file: " & arrFiles(i)
		end if
	next
	Set objFSO = Nothing
end sub

sub invokeFindPathHijack()
	Const FILE_ADD_FILE = &h000002
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Set objShell = CreateObject("WScript.Shell")
	strPath = objShell.ExpandEnvironmentStrings("%path%")
	set objShell = Nothing
	
	arrPaths = Split(strPath, ";")
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	For i = 0 to ubound(arrPaths)
		'-' If the path ends in a backslash, strip it; the FolderExists() check doesn't like that
		if (Right(arrPaths(i), 1) = "\") then 
			arrPaths(i) = mid(arrPaths(i), 1, Len(arrPaths(i))-1)
		end if

		if objFSO.FolderExists(arrPaths(i)) Then
			Set objShare = objWMIService.Get("Win32_Directory.Name='" & arrPaths(i) & "'")
			
			'-' See if the effective permissions say we have write permissions
			isWritable = objShare.GetEffectivePermission(FILE_ADD_FILE)
			if isWritable then 
				wscript.echo "[+] Hijackable .dll path: " & arrPaths(i)
			end if
		Else
			Wscript.Echo "[+] Path does not exist - " & arrPaths(i)
		End if
	next
	Set objShare = Nothing
	Set colListOfServices = Nothing
	Set objWMIService = Nothing
end sub

sub getRegAlwaysInstallElevated
	on error resume next
	Set objShell = CreateObject("Wscript.Shell")
	instValue = objShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\Installer\")
	if err.number = 0 then 
		LMAIEvalue = objShell.RegRead("HKLM\SOFTWARE\Policies\Microsoft\Windows\Installer\AlwaysInstallElevated")
		if err.number = 0 and LMAIEvalue <> 0 then 
			CUAIEvalue = objShell.RegRead("HKCU\SOFTWARE\Policies\Microsoft\Windows\Installer\AlwaysInstallElevated")
			if err.number = 0 and CUAIEvalue <> 0 then 
				wscript.echo "AlwaysInstallElevated enabled on this machine!"
			else 
				wscript.echo "AlwaysInstallElevated not enabled on this machine."
			end if
		else 
			wscript.echo "AlwaysInstallElevated not enabled on this machine."
		end if
	end if
	Set objShell = Nothing
	on error goto 0
end sub

sub checkAutoAdminLogon()
	on error resume next

	Set objShell = CreateObject("Wscript.Shell")
	AALvalue = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AutoAdminLogon")
	if err.number = 0 and AALvalue <> 0 then 
		
		defaultDomainName = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultDomainName")
        defaultUserName = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultUserName")
        defaultPassword = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultPassword")
		
		if NOT isEmpty(defaultUserName) Then
			Wscript.Echo "[+] Autologon default credentials: " & defaultDomainName & ", " & defaultUserName & ", " & defaultPassword
		end if

		altDefaultDomainName = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AltDefaultDomainName")
        altDefaultUserName = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AltDefaultUserName")
        altDefaultPassword = objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\AltDefaultPassword")
		
		if NOT isEmpty(altDefaultUserName) Then
			Wscript.Echo "[+] Autologon alt credentials: " & altDefaultDomainName & ", " & altDefaultUserName & ", " & altDefaultPassword
		end if
		
	end if
	Set objShell = Nothing
	on error goto 0
end sub
