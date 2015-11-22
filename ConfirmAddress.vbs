option explicit

'Canada Post - Postal Code Lookup URL
'curl "https://ws1.postescanada-canadapost.ca/AddressComplete/Interactive/Find/v2.10/json3ex.ws?Key=ea98-jc42-tf94-jk98&Country=CAN&SearchTerm=2956"%"2029th"%"20a&LanguagePreference=en&LastId=&SearchFor=Everything&OrderBy=UserLocation&$block=true&$cache=true&MaxSuggestions=7&MaxResults=100" -H "Origin: https://www.canadapost.ca" -H "Accept-Encoding: gzip, deflate, sdch" -H "Accept-Language: en-US,en;q=0.8" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.86 Safari/537.36" -H "Accept: */*" -H "Referer: https://www.canadapost.ca/cpo/mc/personal/postalcode/fpc.jsf" -H "Connection: keep-alive" --compressed

dim colNamedArguments: Set colNamedArguments = WScript.Arguments.Named

dim sAddress 'as string
dim sURL' as string
dim sReturned 'as string
dim sInputAddress ' as string

dim bVerbose ' as boolean
dim sVerbose ' as string
dim bSilent ' as boolean
dim sSilent ' as boolean

dim bClearedToProceed ' as boolean
dim bCurlVersionOK ' as boolean
dim bParametersOK ' as boolean

dim aFieldWidths ' as array

dim fso ' as file scripting object
dim sInputType ' as string
dim sInputFile ' as string
dim oInputFile ' as File
dim sOutputType ' as string
dim sOutputFile ' as string
dim oOutputFile ' as File

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
		if .Exists("IF") = -1 and .Exists("IA") = 0 then 
			sInputFile = .Item("IF")
			sInputType = "File"
		elseif .Exists("IF") = 0 and .Exists("IA") = -1 then 
			sInputAddress = .Item("IA")
			sInputType = "SingleAddress"
		end if
		if .Exists("OF") = -1 then
			sOutputFile = .Item("OF")
			sOutputType = "File"		
		end if
		sVerbose = .Item("V")
		sSilent = .Item("S")
		end with ' the colNamedArguments one
	
	if sVerbose = "True" then
		bVerbose = 1
	else 
		bVerbose = 0
	end if

	if sSilent = "True" then
		bSilent = 1
	else 
		bSilent = 0
	end if

	if bSilent = 0 then
		call OutputUsage
		call OutputNotes
	end if

	if sInputType = "File" or sOutputType = "File" then
		set fso = CreateObject("Scripting.FileSystemObject")
		if sInputType = "File" then
			set oInputFile = fso.OpenTextFile(sInputFile, 1)
		end if
		if sOutputType = "File" then
			set oOutputFile = fso.CreateTextFile(sOutputFile,True)
		end if	
	end if
				
	redim aFieldWidths(4)
	aFieldWidths(0) = 51
	aFieldWidths(1) = 8
	aFieldWidths(2) = 12
	aFieldWidths(3) = 18
	aFieldWidths(4) = 51

	'wscript.echo "Query Type:  " & sInputType
	
	if bVerbose = 1 then 
		call OutputHeader(aFieldWidths)
	end if

	if sInputType = "SingleAddress" then
		'URL Encode the passed in address
		sURL = BuildCanadaPostURL(sInputAddress)
		sReturned = GetResultFromURL(sURL)
		'wscript.echo sReturned
		call ProcessSingleAddress(sInputAddress, sReturned, aFieldWidths, oOutputFile, sOutputType, bVerbose)
	else
		
		'These variable should be renamed and the dim moved to the top
		dim sFileRow ' as string
		dim iFileRow: iFileRow = 0 ' integer
		
		if bSilent = 0 and bVerbose = 0 then
			wscript.stdout.write "Processing Line: "
		end if 
		Do While oInputFile.AtEndOfStream <> True
			if bSilent = 0 and bVerbose = 0 then
				iFileRow = iFileRow + 1
				if iFileRow <> 1 then
					wscript.stdout.write string(len(iFileRow), chr(8)) ' backspace Character 	
				end if
				wscript.stdout.write iFileRow 
			end if 
			sFileRow = oInputFile.ReadLine
			'this if is just to allow blank rows in the source file for testing purposes
			if left(sFileRow, 7) = "Comment"  then 
				if bVerbose = 1 then
					wscript.echo sFileRow
				end if
			elseif sFileRow <> "" then
				sURL = BuildCanadaPostURL(sFileRow)
				sReturned = GetResultFromURL(sURL)
				call ProcessSingleAddress(sFileRow, sReturned, aFieldWidths, oOutputFile, sOutputType, bVerbose)
			else 
				if bVerbose = 1 then
					wscript.echo ""
				end if
			end if
		Loop
		oInputFile.Close	
	end if
	
end if

sub OutputRowToFile(aResults, oOutputFile)
	
	dim sDBLQuoteCode: sDBLQuoteCode = chr(34)'as string
	dim iField ' as integer
	dim sOutput 'as string
	
	for iField = 0 to 4
		sOutput = sOutput & sDBLQuoteCode & aResults(iField) & sDBLQuoteCode & ";" 
		' Random delimiter may have to trim it off the end
	next	
	oOutputFile.writeline(sOutput)

end sub

sub ProcessSingleAddress(byval sSearchTerm, byval sResult, byval aFieldWidths, byval oOutputFile, byval sOutputType, byval sVerbose)
	
	dim iContainerCount ' as integer
	dim sID ' as string
	dim sCanPostText ' as string
	dim sOutputLine ' as string
	dim iResultCode ' as integer
	dim aResults ' as array
	
	iContainerCount = RetrieveCanadaPostParameter(sResult, "ContainerCount")
	sID = RetrieveCanadaPostParameter(sResult, "Id")
	sCanPostText = RetrieveCanadaPostParameter(sResult, "Text")

	if iContainerCount = 1 and mid(sID, 5, 1) = "B" then
		'wscript.echo sAddress & " - " & sCanPostText
		if sSearchTerm = sCanPostText then
			iResultCode = 2
		else
			iResultCode = 1
		end if
	else
		iResultCode = 0
	end if
	
	if bVerbose = 1 then
		' this will probably have to be bullet proofed against addresses longer that the field lengths
		sOutputLine = sOutputLine & sSearchTerm & string(aFieldWidths(0) - len(sSearchTerm), " ")
		sOutputLine = sOutputLine & iResultCode & string(aFieldWidths(1) - len(iResultCode), " ")
		sOutputLine = sOutputLine & iContainerCount & string(aFieldWidths(2) - len(iContainerCount), " ")
		sOutputLine = sOutputLine & sID & string(aFieldWidths(3) - len(sID), " ")
		if iResultCode <> 0 then 
			sOutputLine = sOutputLine & sCanPostText & string(aFieldWidths(4) - len(sCanPostText), " ")
		else 
			sOutputLine = sOutputLine & left(sCanPostText, aFieldWidths(4) - 4) & "..."
		end if
		wscript.echo sOutputLine
	end if
	
	if sOutputType = "File" then 
		redim aResults(4)
		aResults(0) = sSearchTerm
		aResults(1) = iResultCode
		aResults(2) = iContainerCount
		aResults(3) = sID
		aResults(4) = sCanPostText
		call OutputRowToFile(aResults, oOutputFile)
	end if 
end sub

function CheckParameters(byval colNamedArguments)

	'I could add something here to check for either verbose or output to file
	'as it would be pointless otherwise.  But let's see what happens when the
	'/Silent parameter is added
	with colNamedArguments
		'wscript.echo .Exists("InputFile")
		'wscript.echo .Exists("InputAddress")
		if .Exists("IF") = 0 and .Exists("IA") = 0 then 
			With wscript
				.echo "Error One of the parameters InputFile or InputAddress is required"
				call OutputUsage
				call OutputNotes
				.quit
			end with
		elseif .Exists("IF") = -1 and .Exists("IA") = -1 then 
			With wscript
				.echo "Either the parameter InputFile or the parameter InputAddress is required"
				call OutputUsage
				call OutputNotes
				.quit
			end with	
		end if
		if .Exists("V") = -1 and .Exists("S") = -1 then 
			if .item("V") = "True" and .item("S") = "True" then
				With wscript
					.echo "Error both silent and verbose cannot be selected"
					call OutputUsage
					call OutputNotes
					.quit
				end with
			end if
		end if
	end with
	CheckParameters = True
	
end function

sub OutputHeader(byval aFieldWidths)

	dim sFirstLine ' as string
	dim sSecondLine ' as string
	dim aHeaderText ' as array
	dim iField ' as integer
	dim iTotalWidth ' as integer
	
	redim aHeaderText(4, 1)
	aHeaderText(0, 0) = "Search Address"
	aHeaderText(1, 0) = "Result"
	aHeaderText(2, 0) = "Container"
	aHeaderText(3, 0) = "Canada"
	aHeaderText(4, 0) = "Canada Post Official Address"
	aHeaderText(0, 1) = ""
	aHeaderText(1, 1) = "Code"
	aHeaderText(2, 1) = "Count"
	aHeaderText(3, 1) = "Post Id"
	aHeaderText(4, 1) = ""

	for iField = 0 to 4
		'wscript.echo aFieldWidths(iField) & ", " & len(aHeaderText(iField))
		sFirstLine = sFirstLine & aHeaderText(iField, 0) & string(aFieldWidths(iField) - len(aHeaderText(iField, 0)), " ")
		sSecondLine = sSecondLine & aHeaderText(iField, 1) & string(aFieldWidths(iField) - len(aHeaderText(iField, 1)), " ")
		iTotalWidth = iTotalWidth + aFieldWidths(iField)
	next
	wscript.echo string(iTotalWidth, "=")
	wscript.echo sFirstLine
	wscript.echo sSecondLine
	wscript.echo string(iTotalWidth, "=")
end sub

Sub OutputUsage

	dim iTotalWidth: iTotalWidth = 140
	with wscript
		.echo string(iTotalWidth, "=")
		.echo "	Usage: "
		.echo "	  cscript ConfirmAddress.vbs params"
		.echo "	  params:"
		.echo "	    /IF:FileName.txt                  : Input file (required if /IA not used"
		.echo "	    /IA=""Single Address To Check""     : Just check one address (Required if /IF not used"
		.echo "	    /V:True|False                     : Verbose. optional.  Default is False"
		.echo "	                                        (Verbose optimized for minimum 150 character wide window)"
		.echo "	                                        (Verbose Output has the Canada Post address truncated when result code is 0)"
		.echo "	    /OF:FileName.txt                  : Output File"
		.echo "	                                        (File Output does not have the Canada Post address truncated when result code is 0)"
		.echo	"	                                        (.txt suffix not required)"
		.echo	"	                                        (if file exists it will be overwritten)"
		.echo "	    /S                                : Silent.  Cannot be used with verbose."
		.echo	"	                                        (does not show progress of file processing when /IF is used)"
	end with
		
end sub

Sub OutputNotes

	dim iTotalWidth: iTotalWidth = 140	
	with wscript
		.echo "		"
		.echo "	Notes: "
		.echo "		"
		.echo "	  Result Code 0:  No distinct address found.  Address too poorly formed or not a valid address."
		.echo "	  Result Code 1:  Valid Address.  A single possible address was found.  Not a perfect match to search term."
		.echo "	  Result Code 2:  Valid Address.  A perfect match was found"
		.echo "		"
		.echo "	  Designated multiple dwelling addresses without the suite number return code 0"
		.echo "	  Do not include postal codes with addresses.  They will resolve to 0."
		.echo "		"
		.echo "	  Debug:  If a line in the input file begins with ""comment"" it will be output to the screen in verbose mode"
		.echo string(iTotalWidth, "=")
		.echo "		"
		
		end with 
	
end Sub

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

