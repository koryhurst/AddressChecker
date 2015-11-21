option explicit
function GetResultFromURL(byval sURL)

	dim objShell: Set objShell = WScript.CreateObject("WScript.Shell") ' as object - the command shell
	dim objExec: Set objExec = objShell.Exec(sURL) ' as object - the command to send
	dim sLine 'as string
	dim sReturned 'as string

	Do
		sLine = objExec.StdOut.ReadLine()
		sReturned = sReturned & sLine & vbcrlf
	Loop While Not objExec.Stdout.atEndOfStream

	'wscript.echo sResult
	GetResultFromURL = sReturned

end function
'if the returned result contains the sFailIndicator string it doesn't exist
'if instr(1, sResult, "CAN|1|433|17") > 0 then

function CurlVersionHandlesHTTPS()

	dim sCurlGetVersion: sCurlGetVersion = "curl -V"   'as string
	dim objShell: Set objShell = WScript.CreateObject("WScript.Shell") ' as object - the command shell
	dim objExec: Set objExec = objShell.Exec(sCurlGetVersion) ' as object - the command to send
	dim sCurlVersion 'as string
	dim sLine 'as string
	dim sResult 'as string

	Do
		sLine = objExec.StdOut.ReadLine()
		sCurlVersion = sCurlVersion & sLine & vbcrlf
	Loop While Not objExec.Stdout.atEndOfStream

	if instr(1, sCurlVersion, "https") = 0 then
		with wscript
			.echo "The version of Curl you are using cannot handle HTTPS."
			.echo "You can download the correct version at http://www.confusedbycode.com/curl/"
			.echo "I recommend choosing the zip option and copying the curl executable to this directory"
		end with
		CurlVersionHandlesHTTPS = false
	else
		CurlVersionHandlesHTTPS = true
	end if 

end function