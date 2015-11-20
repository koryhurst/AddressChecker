'Success
'usage cscript.exe ConfirmAddress.vbs "2977 29th Ave E, Vancouver, BC"
'usage cscript.exe ConfirmAddress.vbs "2977 29th Ave East, Vancouver, BC"
'Fail (illustrating a real invalid address)
'usage cscript.exe ConfirmAddress.vbs "2958 29th Ave East, Vancouver, BC"

'Canada Post - Postal Code Lookup URL
'curl "https://ws1.postescanada-canadapost.ca/AddressComplete/Interactive/Find/v2.10/json3ex.ws?Key=ea98-jc42-tf94-jk98&Country=CAN&SearchTerm=2956"%"2029th"%"20a&LanguagePreference=en&LastId=&SearchFor=Everything&OrderBy=UserLocation&$block=true&$cache=true&MaxSuggestions=7&MaxResults=100" -H "Origin: https://www.canadapost.ca" -H "Accept-Encoding: gzip, deflate, sdch" -H "Accept-Language: en-US,en;q=0.8" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.86 Safari/537.36" -H "Accept: */*" -H "Referer: https://www.canadapost.ca/cpo/mc/personal/postalcode/fpc.jsf" -H "Connection: keep-alive" --compressed

dim objShell ' as object - the command shell
dim objExec ' as object - the command to send
dim sAddress 'as string
dim sURLEncAddress 'as string
dim sURLPrefix ' as string
dim sURLSuffix ' as string
dim sURL' as string
dim sDBLQuoteCode 'as string
dim sLine 'as string
dim sReturned 'as string
dim sFailIndicator 'as string
dim bAddressExists 'as boolean
dim iStart 'as integer
dim iFinish 'as integer
dim sSearchForAddressStart ' as string
dim sCanPostAddress 'as string
dim sCurlGetVersion 'as string
dim sCurlVersion 'as string

sDBLQuoteCode = chr(34)
sFailIndicator = "CAN|1|433|17"
sSearchForAddressStart = "Text"

sAddress = WScript.Arguments(0) 
if sAddress = "" then
	wscript.echo "Usage cscript ConfirmAddress ""Address To Check"""
else
	sURLEncAddress = replace(sAddress, " ", sDBLQuoteCode & "%" & sDBLQuoteCode & "20")
	'wscript.echo sURLEncAddress

	sURLPrefix = "curl " & sDBLQuoteCode & "https://ws1.postescanada-canadapost.ca/AddressComplete/Interactive/Find/v2.10/json3ex.ws?Key=ea98-jc42-tf94-jk98&Country=CAN&SearchTerm="
	'wscript.echo sURLPrefix

	'suffix with double quote substitution
	sURLSuffix = "&LanguagePreference=en&LastId=&SearchFor=Everything&OrderBy=UserLocation&$block=true&$cache=true&MaxSuggestions=7&MaxResults=100" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Origin: https://www.canadapost.ca" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Accept-Encoding: gzip, deflate, sdch" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Accept-Language: en-US,en;q=0.8" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.86 Safari/537.36" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Accept: */*" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Referer: https://www.canadapost.ca/cpo/mc/personal/postalcode/fpc.jsf" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Connection: keep-alive" & sDBLQuoteCode & " --compressed"
	'wscript.echo sURLSuffix

	sURL = sURLPrefix & sURLEncAddress & sURLSuffix
	'wscript.echo sURL
	sCurlGetVersion = "curl -V" 'as string

	Set objShell = WScript.CreateObject("WScript.Shell")

	Set objExec = objShell.Exec(sCurlGetVersion)

	Do
		sLine = objExec.StdOut.ReadLine()
		sCurlVersion = sCurlVersion & sLine & vbcrlf
	Loop While Not objExec.Stdout.atEndOfStream

	if instr(1, sCurlVersion, "https") = 0 then
		wscript.echo "The version of Curl you are using cannot handle HTTPS."
		wscript.echo "You can download the correct version at http://www.confusedbycode.com/curl/"
		wscript.echo "I recommend choosing the zip option and copying the curl executable to this directory"
	else
		'bad practice - reusing variable
		set objExec = nothing
		Set objExec = objShell.Exec(sURL)

		Do
			'bad practice - reusing variable
			sLine = objExec.StdOut.ReadLine()
			sResult = sResult & sLine & vbcrlf
		Loop While Not objExec.Stdout.atEndOfStream

		'wscript.echo sAddress
		'wscript.echo sResult


		'if the returned result contains the sFailIndicator string it doesn't exist
		if instr(1, sResult, "CAN|1|433|17") > 0 then
			bAddressExists = False
			wscript.echo "Address is Not Valid"
		else 
			bAddressExists = True
			iStart = instr(1, sResult, "Text") + 7
			'wscript.echo iStart
			iFinish = instr(iStart, sResult, sDBLQuoteCode)
			'wscript.echo iFinish
			sCanPostAddress = mid(sResult, iStart, iFinish - iStart)
			wscript.echo "Searched for Address        :  " & sAddress
			wscript.echo "Canada Post Official Address:  " & sCanPostAddress
			wscript.echo "Result                      :  Address is Valid"
		end if


		set objExec = nothing
		set objShell = nothing


	end if 'the if for curl version check

end if 'the if for parameter check
