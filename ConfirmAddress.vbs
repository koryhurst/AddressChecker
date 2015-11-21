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
	wscript.echo sReturned
	ProcessResult(sReturned)
end if


wscript.quit

function ProcessResult(byval sResult)
	
	'sFailIndicator = "CAN|1|433|17"
	dim sFailIndicator: sFailIndicator = "CAN|1"'as string
	dim sDBLQuoteCode: sDBLQuoteCode = chr(34) 'as string
	dim sSearchForAddressStart: sSearchForAddressStart = "Text" ' as string
	dim bAddressExists 'as boolean
	dim iStart 'as integer
	dim iFinish 'as integer
	dim sCanPostAddress 'as string
	
	if instr(1, sResult, sFailIndicator) > 0 then
		bAddressExists = False
		wscript.echo "No Address Found"
	else 
		bAddressExists = True
		'extract the returned address
		iStart = instr(1, sResult, "Text") + 7
		iFinish = instr(iStart, sResult, sDBLQuoteCode)
		sCanPostAddress = mid(sResult, iStart, iFinish - iStart)
		with wscript
			.echo "Searched for Address        :  " & sAddress
			.echo "Canada Post Address Found   :  " & sCanPostAddress
			if mid(sResult, 19, 1) = 1 then
				.echo "Result                      :  A single possible address was found"
			else
				.echo "Result                      :  " & mid(sResult, 19, 1) & " Possible results found address listed is best guess"
			end if 
		end with 
	end if

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

