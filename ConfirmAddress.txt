option explicit
'Success
'usage cscript.exe ConfirmAddress.vbs "2977 29th Ave E, Vancouver, BC"
'usage cscript.exe ConfirmAddress.vbs "2977 29th Ave East, Vancouver, BC"
'Fail (illustrating a real invalid address)
'usage cscript.exe ConfirmAddress.vbs "2958 29th Ave East, Vancouver, BC"

'Canada Post - Postal Code Lookup URL
'curl "https://ws1.postescanada-canadapost.ca/AddressComplete/Interactive/Find/v2.10/json3ex.ws?Key=ea98-jc42-tf94-jk98&Country=CAN&SearchTerm=2956"%"2029th"%"20a&LanguagePreference=en&LastId=&SearchFor=Everything&OrderBy=UserLocation&$block=true&$cache=true&MaxSuggestions=7&MaxResults=100" -H "Origin: https://www.canadapost.ca" -H "Accept-Encoding: gzip, deflate, sdch" -H "Accept-Language: en-US,en;q=0.8" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.86 Safari/537.36" -H "Accept: */*" -H "Referer: https://www.canadapost.ca/cpo/mc/personal/postalcode/fpc.jsf" -H "Connection: keep-alive" --compressed

dim sAddress 'as string
dim sURL' as string
dim sReturned 'as string

'on error resume next 
call Include("CurlFunctions")

sAddress = WScript.Arguments(0) 
if sAddress = "" then
	wscript.echo "Usage: cscript ConfirmAddress.vbs ""Address To Check"""
	wscript.quit
end if

if CurlVersionHandlesHTTPS = True then 
	'URL Encode the passed in address
	sURL = BuildCanadaPostURL(sAddress)
	sReturned = GetResultFromURL(sURL)
	'wscript.echo sReturned
	ProcessResult(sReturned)
end if



wscript.quit

function ProcessResult(byval sResult)
	
	dim iContainerCount ' as integer
	dim sID ' as string
	dim sCanPostText ' as string
	
	iContainerCount = RetrieveCanadaPostParameter(sResult, "ContainerCount")
	sID = RetrieveCanadaPostParameter(sResult, "Id")
	sCanPostText = RetrieveCanadaPostParameter(sResult, "Text")
	'if iContainerCount = 1 and isnumeric(right(sID, 7)) then
		with wscript
			.echo "Searched for Address         :  " & sAddress
			.echo "Canada Post Container Count  :  " & iContainerCount
			.echo "Canada Post ID               :  " & sID
			.echo "Canada Post Text             :  " & sCanPostText
			if iContainerCount = 1 and isnumeric(right(sID, 7)) then
				.echo "Canada Post Address          :  " & sCanPostText
				if sAddress = sCanPostText then
					.echo "Result                       :  Valid Address.  A perfect match was found"
				else
					.echo "Result                       :  Valid Address.  A single possible address was found.  Not a perfect match to search term."
				end if
			else
				.echo "Result                       :  No distinct address found.  Sought after address either too poorly formed or not a valid address"
			end if
		end with 

end function
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

