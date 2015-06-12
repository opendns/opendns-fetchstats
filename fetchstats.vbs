' OpenDNS Stats Fetch for Windows
' Based on original fetchstats script from Richard Crowley
' Brian Hartvigsen <brian.hartvigsen@opendns.com>
'
' Usage: cscript /NoLogo fetchstats.vbs <username> <network-id> <YYYY-MM-DD> [<YYYY-MM-DD>]

' Instantiate object here so that cookies will be tracked.
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

Function GetUrlData(strUrl, strMethod, strData)
	objHTTP.Open strMethod, strUrl
	If strMethod = "POST" Then
		objHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
	End If
	objHTTP.Option(6) = False
	objHTTP.Send(strData)

	If Err.Number <> 0 Then
		GetUrlData = "ERROR: " & Err.Description & vbCrLf & Err.Source & " (" & Err.Nmber & ")"
	Else
		GetUrlData = objHTTP.ResponseText
	End If
End Function

LOGINURL="https://login.opendns.com/?source=dashboard"
CSVURL="https://dashboard.opendns.com"

Sub Usage()
	Wscript.StdErr.Write "Usage: <username> <network_id> <YYYY-MM-DD> [<YYYY-MM-DD>]" & vbCrLf
	WScript.Quit 1
End Sub

If Wscript.Arguments.Count < 3 Or Wscript.Arguments.Count > 4 Then
	Usage
End If

Username = Wscript.Arguments.Item(0)
If Len(Username) = 0 Then: Usage
Network = Wscript.Arguments.Item(1)
If Len(Network) = 0 Then: Usage

Set regDate = new RegExp
regDate.Pattern = "^\d{4}-\d{2}-\d{2}$"

DateRange = Wscript.Arguments.Item(2)
If Not regDate.Test(DateRange) Then: Usage

If Wscript.Arguments.Count = 4 Then
	ToDate = Wscript.Arguments.Item(3)
	If Not regDate.Test(toDate) Then: Usage
	DateRange = DateRange & "to" & ToDate
End If

WScript.StdErr.Write "Password: "

' Are they running Vista or 7?
On Error Resume Next
Set objPassword = CreateObject("ScriptPW.Password") 
If objPassword <> Null Then
    Password = objPassword.GetPassword()
Else
    Password = WScript.StdIn.ReadLine
End If
Wscript.StdErr.Write vbCrLf
On Error GoTo 0

Set regEx = New RegExp
regEx.IgnoreCase = true
regEx.Pattern = ".*name=""formtoken"" value=""([0-9a-f]*)"".*"

data = GetUrlData(LOGINURL, "GET", "")

Set Matches = regEx.Execute(data)
token = Matches(0).SubMatches(0)

data = GetUrlData(LOGINURL, "POST", "formtoken=" & token & "&username=" & replace(escape(Username),"+","%2B")  & "&password=" & replace(escape(Password),"+","%2B")  & "&sign_in_submit=foo")

regEx.Pattern = ".*Logging you in.*"
Set loginMatches = regEx.Execute(data)

If loginMatches.Count = 0 Then
	Wscript.StdErr.Write "Login Failed. Check username and password" & vbCrLf
	WScript.Quit 1
End If

page=1
Do While True
	data = GetUrlData(CSVURL & "/stats/" & Network & "/topdomains/" & DateRange & "/page" & page & ".csv", "GET", "")
	If page = 1 Then
		If LenB(data) = 0 Then
			WScript.StdErr.Write "You can not access " & Network & vbCrLf
			WScript.Quit 2
		ElseIf InStr(data, "<!DOCTYPE") Then
		    Wscript.StdErr.Write "Error retrieving data.  Date range may be outside of available data."
		    Wscript.Quit 2
		End If
	Else
		' First line is always header
		data=Mid(data,InStr(data,vbLf)+1)
	End If
	If LenB(Trim(data)) = 0 Then
		Exit Do
	End If
	Wscript.StdOut.Write data
	page = page + 1
Loop
