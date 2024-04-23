<%

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

'This section pulls the user information to be used to display a hello message.
Set objRs=objConn.Execute("Select * from UserInfo WHERE ID="& Session("UID") & "")
Set objRs2=objConn.Execute("Select * from Domains WHERE Country = '"& objRs("Country") &"'")
	strSend =objRs("Email") & objRs2("Domain")
' Create an instance of the NewMail object.
Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
    
' Set the properties of the object
objCDOMail.From = "support@purefishing.com"
objCDOMail.To = strSend
objCDOMail.Subject = "Online Help Desk Password"
objCDOMail.Body = "Your recent request for your password has been processed." & vbcrlf & "Your password is " & objRs("Password")

objCDOMail.Send

objRs.close
objConn.Close
Set objRs = Nothing
Set objConn = Nothing

Set objCDOMail = Nothing

Session("UID")=""

'Response.Redirect "login.asp"
%>


<script>

function gotologin(){
alert("Your password has been E-mailed to you.")
window.location="login.asp"
}

</script>

<html>
<title>O.H.D. Password Lookup</title>
<body onload="gotologin()">
</body>
</html>