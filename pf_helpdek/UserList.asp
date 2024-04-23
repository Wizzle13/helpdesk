<!--#include file="_head.asp"-->
<!--
Programer: Chris Burton    
Date Started: 7-9-01

Description:
This page checks the  database for users First and Last name
then uses it to display a hello message to the user.
-->
<%
Response.Write "</td>"
Response.Write "<td valign=Top>"
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

Logon=Request.Form("Logon_Name") + Request.Form("Domain")

'This section pulls the user information to be used to display a hello message.
Set objRs=objConn.Execute("Select * from UserInfo ORDER by LastName, FirstName")


Response.Write "<table border=0 cellpadding=0 cellborder=0>"
Response.Write "<tr>"
Response.Write "<td><B>User Name</td>"
Response.Write "</tr>"		
	While Not objRs.EOF
		Response.Write "<tr>"
		Response.Write "<td><a href=Add.asp?Num=" & objRs("ID") &">" & objRs("LastName") &", " & objRs("FirstName") &"</a></td>"	
		Response.Write "</tr>"	
		objRs.MoveNext
	Wend

	Response.Write "</table>"
objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing


	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

%>
</body>
</html>
