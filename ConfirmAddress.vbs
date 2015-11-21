option explicit
'Success
'usage cscript.exe ConfirmAddress.vbs "2977 29th Ave E, Vancouver, BC"
'usage cscript.exe ConfirmAddress.vbs "2977 29th Ave East, Vancouver, BC"
'Fail (illustrating a real invalid address)
'usage cscript.exe ConfirmAddress.vbs "2958 29th Ave East, Vancouver, BC"

'Canada Post - Postal Code Lookup URL
'curl "https://ws1.postescanada-canadapost.ca/AddressComplete/Interactive/Find/v2.10/json3ex.ws?Key=ea98-jc42-tf94-jk98&Country=CAN&SearchTerm=2956"%"2029th"%"20a&LanguagePreference=en&LastId=&SearchFor=Everything&OrderBy=UserLocation&$block=true&$cache=true&MaxSuggestions=7&MaxResults=100" -H "Origin: https://www.canadapost.ca" -H "Accept-Encoding: gzip, deflate, sdch" -H "Accept-Language: en-US,en;q=0.8" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.86 Safari/537.36" -H "Accept: */*" -H "Referer: https://www.canadapost.ca/cpo/mc/personal/postalcode/fpc.jsf" -H "Connection: keep-alive" --compressed

dim colNamedArguments: Set colNamedArguments = WScript.Arguments.Named
dim sAddress 'as string
dim sURL' as string
dim sReturned 'as string
dim sInputFile ' as string
dim sInputAddress ' as string
dim sInputType ' as string
dim bVerbose ' as boolean
dim sVerbose ' as string
dim bClearedToProceed ' as boolean
dim bCurlVersionOK ' as boolean
dim bParametersOK ' as boolean

'on error resume next 
call Include("CurlFunctions")

'This whole section should be functionalized
'at least the check to have the right parameters
'Then back here at main I can assign them if bClearedToProceed is true
bClearedToProceed = False
'this one could be a sub, as it just bails if it encounters trouble
bParametersOK = CheckParameters(colNamedArguments)
bCurlVersionOK = CurlVersionHandlesHTTPS
bClearedToProceed = bParametersOK and bCurlVersionOK

if bClearedToProceed = True then 
	with colNamedArguments
		'wscript.echo .Exists("InputFile")
		'wscript.echo .Exists("InputAddress")
		if .Exists("InputFile") = -1 and .Exists("InputAddress") = 0 then 
			sInputFile = .Item("InputFile")
			sInputType = "File"
		elseif .Exists("InputFile") = 0 and .Exists("InputAddress") = -1 then 
			sInputAddress = .Item("InputAddress")
			sInputType = "SingleAddress"
		end if
		sVerbose = .Item("Verbose")
	end with ' the colNamedArguments one
	
	if sVerbose = "True" then
		bVerbose = 1
	else 
		bVerbose = 0
	end if
	
	if bVerbose = 1 then 
		call OutputNotes
	end if
	
	if sInputType = "SingleAddress" then
		'URL Encode the passed in address
		sURL = BuildCanadaPostURL(sInputAddress)
		sReturned = GetResultFromURL(sURL)
		'wscript.echo sReturned
		call ProcessSingleAddress(sInputAddress, sReturned)
	else
		dim fso: set fso = CreateObject("Scripting.FileSystemObject")
'		dim sFullFileName:  sFullFileName = fso.BuildPath(CurrentDirectory, sInputFile)
		dim oFile: set oFile = fso.OpenTextFile(sInputFile, 1)
		dim sFileRow ' as string
		Do While oFile.AtEndOfStream <> True
			sFileRow = oFile.ReadLine
			sURL = BuildCanadaPostURL(sFileRow)
			sReturned = GetResultFromURL(sURL)
			call ProcessSingleAddress(sFileRow, sReturned)
		Loop
		oFile.Close	
	end if
	
end if

wscript.quit

function CheckParameters(byval colNamedArguments)

	with colNamedArguments
		'wscript.echo .Exists("InputFile")
		'wscript.echo .Exists("InputAddress")
		if .Exists("InputFile") = 0 and .Exists("InputAddress") = 0 then 
			With wscript
				.echo "Error One of the parameters InputFile or InputAddress is required"
				call OutputUsage
				.quit
			end with
		elseif .Exists("InputFile") = -1 and .Exists("InputAddress") = -1 then 
			With wscript
				.echo "Either the parameter InputFile or the parameter InputAddress is required"
				call OutputUsage
				.quit
			end with	
		end if
	end with
	CheckParameters = True
	
end function

Sub OutputUsage
		
		with wscript
			.echo "Usage: "
			.echo "  cscript ConfirmAddress.vbs params"
			.echo "  params:"
			.echo "  /InputFile:FileName.txt or /InputAddress=""Single Address To Check"" ONE REQUIRED"
			.echo "  /Verbose:True|False  optional.  Default is False"
		end with
		call OutputNotes
		
end sub
Sub OutputNotes
	
	with wscript
		.echo "Notes: "
		.echo "Result Code 2:  Valid Address.  A perfect match was found"
		.echo "Result Code 1:  Valid Address.  A single possible address was found.  Not a perfect match to search term."
		.echo "Result Code 0:  No distinct address found.  Address too poorly formed or not a valid address.  "
		.echo " "
		.echo "**Multiple dwelling addresses without the suite number return code 0"
		.echo " "
	end with 
	
end Sub
sub ProcessSingleAddress(byval sSearchTerm, byval sResult)
	
	dim iContainerCount ' as integer
	dim sID ' as string
	dim sCanPostText ' as string
	
	iContainerCount = RetrieveCanadaPostParameter(sResult, "ContainerCount")
	sID = RetrieveCanadaPostParameter(sResult, "Id")
	sCanPostText = RetrieveCanadaPostParameter(sResult, "Text")
	'this output now has to be columnized
	'if iContainerCount = 1 and isnumeric(right(sID, 7)) then
	with wscript
		.echo "Searched for Address         :  " & sSearchTerm
		.echo "Canada Post Container Count  :  " & iContainerCount
		.echo "Canada Post ID               :  " & sID
		.echo "Canada Post Text             :  " & sCanPostText
'			if iContainerCount = 1 and isnumeric(right(sID, 7) then
		if iContainerCount = 1 and mid(sID, 5, 1) = "B" then
			.echo "Canada Post Address          :  " & sCanPostText
			if sAddress = sCanPostText then
				.echo "Result Code                  :  2"
			else
				.echo "Result Code                  :  1"
			end if
		else
				.echo "Result Code                  :  0"
		end if
		.echo " " 'spacer
	end with 

end sub

function RetrieveCanadaPostParameter(byval sResultSet, byval sParameterName)

	dim sDBLQuoteCode: sDBLQuoteCode = chr(34) 'as string
	dim sReturnValue ' as string
	dim iStart 'as integer
	dim iFinish 'as integer
	'{"ContainerCount":1,"Items":[{"Id":"CAN|8|833|1897","Text":"Sussex Dr, Ottawa, ON","Highlight":"","Cursor":0,"Description":"55 Results","Next":"Find"}]}
	'{"ContainerCount":1,"Items":[{"Id":"CAN|B|3730784","Text":"2961 29th Ave E, Vancouver, BC","Highlight":"","Cursor":0,"Description":"","Next":"Retrieve"}]}
	if sParameterName = "ContainerCount" then
		sReturnValue = mid(sResultSet, 19, 1)
	elseif sParameterName = "Id" then
		iStart = instr(1, sResultSet, "Id") + 5
		iFinish = instr(iStart, sResultSet, sDBLQuoteCode)
		sReturnValue = mid(sResultSet, iStart, iFinish - iStart)
	elseif sParameterName = "Text" then
		iStart = instr(1, sResultSet, "Text") + 7
		iFinish = instr(iStart, sResultSet, sDBLQuoteCode)
		sReturnValue = mid(sResultSet, iStart, iFinish - iStart)
	end if
	RetrieveCanadaPostParameter = sReturnValue
	
end function

function BuildCanadaPostURL(byval sAddress)

	dim sDBLQuoteCode: sDBLQuoteCode = chr(34)'as string
	dim sURLEncAddress 'as string
	dim sURLPrefix ' as string
	dim sURLSuffix ' as string
	dim sBuiltUrl 'as string
	
	'sDBLQuoteCode = chr(34)
	
	sURLEncAddress = replace(sAddress, " ", sDBLQuoteCode & "%" & sDBLQuoteCode & "20")
	sURLPrefix = "curl " & sDBLQuoteCode & "https://ws1.postescanada-canadapost.ca/AddressComplete/Interactive/Find/v2.10/json3ex.ws?Key=ea98-jc42-tf94-jk98&Country=CAN&SearchTerm="
	sURLSuffix = "&LanguagePreference=en&LastId=&SearchFor=Everything&OrderBy=UserLocation&$block=true&$cache=true&MaxSuggestions=7&MaxResults=100" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Origin: https://www.canadapost.ca" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Accept-Encoding: gzip, deflate, sdch" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Accept-Language: en-US,en;q=0.8" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.86 Safari/537.36" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Accept: */*" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Referer: https://www.canadapost.ca/cpo/mc/personal/postalcode/fpc.jsf" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Connection: keep-alive" & sDBLQuoteCode & " --compressed"
	sBuiltUrl = sURLPrefix & sURLEncAddress & sURLSuffix
	'wscript.echo sBuiltUrl
	BuildCanadaPostURL = sBuiltUrl

end function

Sub Include(sFileName)

  dim fso: set fso = CreateObject("Scripting.FileSystemObject")
	dim oFile: set oFile = fso.OpenTextFile(sFileName & ".vbs", 1)
	dim sFileContents ' as string
  sFileContents = oFile.ReadAll
  oFile.Close
  ExecuteGlobal sFileContents

End Sub

