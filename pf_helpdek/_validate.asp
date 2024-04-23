<%
Dim DCount
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
Domain=Request.Form("Domain")
Logon=Request.Form("Logon_Name")
'This section pulls the user information to be used to display a hello message.

Set objREC1=objConn.Execute("Select Country from Domains WHERE Domain = '" & Domain & "'")
Set DomainCount=objConn.Execute("Select Count(*) from Domains WHERE Domain = '" & Domain & "'")
DCount=1
While DCount <= DomainCount(0)

	Country=objRec1("Country")
	Set objREC=objConn.Execute("Select * from UserInfo")
	Set rsCount=objConn.Execute("SELECT Count(*) from UserInfo WHERE Email = '" & Logon & "' and Country= '" & Country &"'")
	IF rsCount(0) > 0 then
		DCount = DomainCount(0)
	End if
	DCount=Dcount + 1	
	objRec1.MoveNext

Wend



'Set Cookie info.
If Request.Form("CookieCheck") = "on" Then
	Response.buffer = True 
	' first we put the data from the form into a variable

	'then we can set the cookie value using this:
	Response.Cookies("Logon") = Logon
	Response.Cookies("Logon").Expires = Date + 365
	Response.Cookies("Domain") = Domain
	Response.Cookies("Domain").Expires = Date + 365
Else
	Response.buffer = False
	Response.Cookies("Logon").Expires = Date - 365
	Response.Cookies("Domain").Expires = Date - 365
End If


' This section checks to see if  the user exists in the Database
If rsCount(0) < 1  then
Response.Write "This username does not exist in our database.  If you think this is incorrect, please contact the IT Help Desk. If you have forgotten your password, click <a href=passwordlookup.asp>here</a>."

objRec.close
objRec1.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing
Set objRec1 = Nothing
Else


' This section gets the users Information
Set objRs=objConn.Execute("Select * from UserInfo WHERE Email = '" & Logon & "' and Country= '" & Country &"'")

' Checks to see if there is a password entered.
If Request.Form("Password") = "" Then
	Session("UID") = objRs("ID")
	
	Response.Redirect "sendpassword.asp"
End If
' Checks to see if the password is correct then sets up the session vairables.
If Request.Form("Password") = objRs("Password") then
	Session("UID") = objRs("ID")
	Session("IsValid") = "True"
	Session("FirstName") = objRs("FirstName")
	Session("LastName") = objRs("LastName")
	Session("Extention") = objRs("Extention")
	Session("Email") = objRs("Email")
	Session("Status") = objRs("Status")
	Session("WorkID") = objRs("WorkerID")
	Session("Password") = objRs("Password")
	Session("Department") = objRs("Department")
	Session("Country") = objRs("Country")
	Session("UserAdded") = objRs("UserAdded")
	objRs.close
	objConn.Close
	Set objRs = Nothing
	Set objConn = Nothing

	%>
	<!--#include file="_WhosOn.Asp"-->
	<%
	If Session("URL") = "/index.asp?" Then
		Session("URL") = ""
	End If
	If Session("UserAdded")="1" then
		Response.Redirect "/Password.asp"
	End if 

	If Session("URL") = "" Then
		Response.Redirect "/main.asp"
	Else
		Response.Redirect Session("URL")
	End If
	
Else
	objRs.close
	objConn.Close
	Set objRs = Nothing
	Set objConn = Nothing

	'Response.Write "<SCRIPT LANGUAGE=JavaScript>"
	'Response.Write "alert('Your password is incorrect, please try again.');"
	'Response.Write "</SCRIPT>"

	Response.Redirect "/login.asp?login=n"
End If

End IF
%>