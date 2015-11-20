'usage cscript.exe ConfirmAddress.vbs "2977 29th Ave East, Vancouver, BC"

'Canada Post - Postal Code Lookup URL
'curl "https://ws1.postescanada-canadapost.ca/AddressComplete/Interactive/Find/v2.10/json3ex.ws?Key=ea98-jc42-tf94-jk98&Country=CAN&SearchTerm=2956"%"2029th"%"20a&LanguagePreference=en&LastId=&SearchFor=Everything&OrderBy=UserLocation&$block=true&$cache=true&MaxSuggestions=7&MaxResults=100" -H "Origin: https://www.canadapost.ca" -H "Accept-Encoding: gzip, deflate, sdch" -H "Accept-Language: en-US,en;q=0.8" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.86 Safari/537.36" -H "Accept: */*" -H "Referer: https://www.canadapost.ca/cpo/mc/personal/postalcode/fpc.jsf" -H "Connection: keep-alive" --compressed>CanadaPostPossibleFail.html 

dim sAddress 'as string
dim sURLEncAddress 'as string
dim sURLPrefix ' as string
dim sURLSuffix ' as string
dim sURL' as string
dim sDBLQuoteCode

sDBLQuoteCode = chr(34)

sAddress = WScript.Arguments(0) 
sURLEncAddress = replace(sAddress, " ", sDBLQuoteCode & "%" & sDBLQuoteCode)

wscript.echo sURLEncAddress
sURLPrefix = "https://ws1.postescanada-canadapost.ca/AddressComplete/Interactive/Find/v2.10/json3ex.ws?Key=ea98-jc42-tf94-jk98&Country=CAN&SearchTerm="
'before fixing the double quotes
'sURLSuffix = "&LanguagePreference=en&LastId=&SearchFor=Everything&OrderBy=UserLocation&$block=true&$cache=true&MaxSuggestions=7&MaxResults=100" -H "Origin: https://www.canadapost.ca" -H "Accept-Encoding: gzip, deflate, sdch" -H "Accept-Language: en-US,en;q=0.8" -H "User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.86 Safari/537.36" -H "Accept: */*" -H "Referer: https://www.canadapost.ca/cpo/mc/personal/postalcode/fpc.jsf" -H "Connection: keep-alive" --compressed>CanadaPostPossibleFail.html"

'after fixing the double quotes
sURLSuffix = "&LanguagePreference=en&LastId=&SearchFor=Everything&OrderBy=UserLocation&$block=true&$cache=true&MaxSuggestions=7&MaxResults=100" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Origin: https://www.canadapost.ca" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Accept-Encoding: gzip, deflate, sdch" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Accept-Language: en-US,en;q=0.8" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.86 Safari/537.36" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Accept: */*" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Referer: https://www.canadapost.ca/cpo/mc/personal/postalcode/fpc.jsf" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Connection: keep-alive" & sDBLQuoteCode & " --compressed>CanadaPostPossibleFail.html"

'after removing the pipe into a text file for testing
sURLSuffix = "&LanguagePreference=en&LastId=&SearchFor=Everything&OrderBy=UserLocation&$block=true&$cache=true&MaxSuggestions=7&MaxResults=100" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Origin: https://www.canadapost.ca" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Accept-Encoding: gzip, deflate, sdch" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Accept-Language: en-US,en;q=0.8" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.86 Safari/537.36" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Accept: */*" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Referer: https://www.canadapost.ca/cpo/mc/personal/postalcode/fpc.jsf" & sDBLQuoteCode & " -H " & sDBLQuoteCode & "Connection: keep-alive" & sDBLQuoteCode & " --compressed"

sURL = sURLPrefix & sAddress & sURLSuffix
wscript.echo sURL

