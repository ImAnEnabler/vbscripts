'-'
'-' Invoke from the command line and it dumps to stdout. Redirect to a file for CSV goodness.
'-'      cscript //nologo getComputerInventoryFromAD.vbs > workstation_inventory.csv
'-' by: @ImAnEnabler

'-' Set the target OU here, where the computers are located you want to inventory
strTarget = "LDAP://OU=Workstations,CN=Computers,DC=yourdomain,DC=local"

'-'  TRY NOT TO EDIT BELOW HERE
'-'  if you do need to add a new column, you'll need to edit three locations below.
'-'  keep the columns ordered properly.

'-' Set up the LDAP AD query
Set objRootDSE = GetObject("LDAP://RootDSE")

' Connect to AD Provider
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

'-' Set up the Command object to perform the query
Set objCmd = CreateObject("ADODB.Command")
Set objCmd.ActiveConnection = objConnection 

'-' generate the query to be run
objCmd.CommandText = "SELECT name, " & _ 
						"description, " & _ 
						"ipHostNumber, " & _ 
						"serialNumber, " & _ 
						"personalPager, " & _ 
						"employeeType, " & _ 
						"operatingSystem, " & _ 
						"operatingSystemServicePack, " & _ 
						"lastLogonTimestamp " & _ 
						"FROM '" & strTarget & "' " & _ 
						"WHERE objectCategory = 'computer' " & _ 
						"ORDER by name" 
						
'-' set up some paramaters for the query
Const ADS_SCOPE_SUBTREE = 2
objCmd.Properties("Page Size") = 100
objCmd.Properties("Timeout") = 30
objCmd.Properties("Searchscope") = ADS_SCOPE_SUBTREE
objCmd.Properties("Cache Results") = False

'-' Execute the query
Set objRecordSet = objCmd.Execute

'-' Iterate through the results
'-' First, make sure we're at the top of the recordset
objRecordSet.MoveFirst
'-' set up the first line as column names
'-' this is all getting dumped to stdout in CSV format
strAllFields = "Name," & _
				"Description," & _
				"ipHostNumber," & _
				"serialNumber," & _
				"personalPager," & _
				"employeeType," & _
				"operatingSystem," & _
				"operatingSystemServicePack," & _
				"lastLogonTimestamp"
'-' output the first line
'-' this is all getting dumped to stdout in CSV format
wscript.Echo strAllFields
				
'-' here's where the loop begins.
'-' Read each record, parse the values into the same order as the column names above.
'-' Some of the fields are of .Type string (202) and some of the fields are .Type variant (12)
'-' the variant types need to be further parsed through with the getValues() function below
'-' the quoteString() function helps with formatting the CSV output
Do Until objRecordSet.EOF
	strAllFields = UCase(objRecordSet.Fields("Name").Value) & "," & _
					quoteString(getValues(objRecordSet.Fields("Description").Value)) & "," & _
					quoteString(getValues(objRecordSet.Fields("ipHostNumber").Value)) & "," & _
					getValues(objRecordSet.Fields("serialNumber").Value) & "," & _
					quoteString(objRecordSet.Fields("personalPager").Value) & "," & _
					quoteString(objRecordSet.Fields("employeeType").Value) & "," & _
					quoteString(objRecordSet.Fields("operatingSystem").Value) & "," & _
					quoteString(objRecordSet.Fields("operatingSystemServicePack").Value) & "," & _
					ConvertDate(objRecordSet.Fields("lastLogonTimestamp").Value)
	'-' then, output the results as a line
	'-' this is all getting dumped to stdout in CSV format
	wscript.Echo strAllFields
	objRecordSet.MoveNext
Loop




'-'  FUNCTIONS BELOW - DO NOT EDIT BELOW THIS LINE

'-' Some of the values returned are not strings, but a collection of strings, 
'-' even if there's only one value.  This function will parse the collection
'-' and return the string value back
function getValues(objField)
	strValue = ""
	'-' if the field is empty, the script breaks, so we check
	if not isNull(objField) Then
		'-' loop through each item in the collection, concatenate into a string
		'-' often, there's only one value, but this is the recommended way to handle it
		for each item in objField
			strValue = strValue & item
		next
	end if
	'-' returns the string
	getValues = strValue
end function

'-' takes a string and adds a "double quote" character to each end
'-' helps with formatting the CSV output
function quoteString(inString)
	quoteString = """" & inString & """"
end function

'-' ripped from the internet and tweaked for our own use
'-' the lastLogonTimestamp field is a 64-bit integer, which correlates to 
'-' the number of 100-nanosecond intervals that have elapsed since 
'-' the 0 hour on January 1, 1601
'-' for this script we're not interested in time, only date
function ConvertDate(objDate)
    ' FUNCTION to convert Integer8 (64-bit) value to a date, adjusted for
    ' time zone bias.
    lngAdjust = 240 '-' this would be -60 for DST
    lngHigh = objDate.HighPart
    lngLow = objDate.LowPart
    ' Account for bug in IADsLargeInteger property methods.
    IF (lngHigh = 0) And (lngLow = 0) THEN
        lngAdjust = 0
    END IF
    lngDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
        + lngLow) / 600000000 - lngAdjust) / 1440
    ConvertDate = DateValue(CDate(lngDate))
end function
