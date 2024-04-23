<!--#include file="_validsession.asp"-->
<!--
Programer: Chris Burton    
Date Started: 7-9-01

Description:
This page checks the  database for users First and Last name
then uses it to display a hello message to the user.

-->
<!--#include file="_head.asp"-->
<%
Response.Write "</td>"
Response.Write "<td valign=Top>"
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

Logon=Request.Form("Logon_Name") + Request.Form("Domain")
Country=session("Country")
'This section pulls the user information to be used to display a hello message.
If Session("Status")="Administrator" then
	If Request("sortby")="FName" then
		Set objRs=objConn.Execute("Select * from UserInfo ORDER by FirstName, LastName")
	End if
	If Request("sortby")="LName" or Request("sortby")=""  then
		Set objRs=objConn.Execute("Select * from UserInfo ORDER by LastName,FirstName")
	End if
	If Request("sortby")="Status"  then
		Set objRs=objConn.Execute("Select * from UserInfo ORDER by Status,FirstName")
	End if
	If Request("sortby")="Department"  then
		Set objRs=objConn.Execute("Select * from UserInfo ORDER by Department,FirstName")
	End if
	If Request("sortby")="Country"  then
		Set objRs=objConn.Execute("Select * from UserInfo ORDER by Country,FirstName")
	End if
	If Request("sortby")="ID"  then
		Set objRs=objConn.Execute("Select * from UserInfo ORDER by ID")
	End if
Else
	If Request("sortby")="FName" then
		Set objRs=objConn.Execute("Select * from UserInfo where Country='"& Country & "' ORDER by FirstName, LastName")
	End if
	If Request("sortby")="LName" or Request("sortby")=""  then
		Set objRs=objConn.Execute("Select * from UserInfo where Country='"& Country & "' ORDER by LastName,FirstName")
	End if
	If Request("sortby")="Status"  then
		Set objRs=objConn.Execute("Select * from UserInfo where Country='"& Country & "' ORDER by Status,FirstName")
	End if
	If Request("sortby")="Department"  then
		Set objRs=objConn.Execute("Select * from UserInfo where Country='"& Country & "' ORDER by Department,FirstName")
	End if
	If Request("sortby")="Country"  then
		Set objRs=objConn.Execute("Select * from UserInfo where Country='"& Country & "' ORDER by Country,FirstName")
	End if
	If Request("sortby")="ID"  then
		Set objRs=objConn.Execute("Select * from UserInfo where Country='"& Country & "' ORDER by ID")
	End if
End if

Response.Write "<table border=0 cellpadding=0 cellborder=0>"
Response.Write "<tr>"
Response.Write "<td><a href=ManageUsers.asp?sortby=LName><B>Last Name</a></td>"
Response.Write "<td><a href=ManageUsers.asp?sortby=FName><B>First Name</a></td>"
Response.Write "<td><a href=ManageUsers.asp?sortby=Status><B>Member Status</a></td>"
Response.Write "<td><a href=ManageUsers.asp?sortby=Department><B>Department</a></td>"
Response.Write "<td><a href=ManageUsers.asp?sortby=Country><B>Country</a></td>"
Response.Write "</tr>"		
	While Not objRs.EOF
		Response.Write "<tr>"
		Response.Write "<td><a href=modify.asp?Num=" & objRs("ID") &">" & objRs("LastName") &" </a></td>"	
		Response.Write "<td><a href=modify.asp?Num=" & objRs("ID") &"> " & objRs("FirstName") &"</a></td>"	
		Response.Write "<td>" & objRs("Status") & "</td>"		
		Response.Write "<td>" & objRs("Department") & "</td>"		
		Response.Write "<td>" & objRs("Country") & "</td>"		
		Response.Write "</tr>"	
		objRs.MoveNext
	Wend
Response.Write "<tr><td><a href=ManageUsers.asp?sortby=ID><B>.</a></td></tr>"
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
