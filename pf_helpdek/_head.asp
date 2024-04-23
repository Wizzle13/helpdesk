<!--#include file="_validsession.asp"-->

<!-- This Page be the top and side for all the pages involved
	with Help Desk v. 2 -->

<html>
<head>



<SCRIPT src="/rightmenu.js" type=text/javascript>

</SCRIPT>

<script>
function logout(){
	var targetURL="http://172.16.3.10/logoff.asp"
	window.location=targetURL
return
}
</script>

<link rel="stylesheet" href="/helpdesk.css" type="text/css">

<%
If Pagename = "ShowTicket" Then
	Response.Write "<title>O.H.D. Ticket #" & Request("Num") & "</title>"
Else
	Response.Write "<title>Online Help Desk</title>"
End If
%>
</head>

<%
If Request.ServerVariables("SERVER_NAME") = "172.16.6.245" Then
	Response.Write "<body background='/images/dev_bg.jpg'>"
Else
	Response.Write "<body>"
End If
'Response.Write Request.ServerVariables("SERVER_NAME")

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

'This section pulls the user information to be used to display a hello message.
Set objRs=objConn.Execute("Select * from UserInfo WHERE ID=" & Session("UID") & "")

	Response.Write "<table border=1 width=600 align=center bgcolor=white>"
	Response.Write "<tr>"
	Response.Write "<td align=center width=125><a href='/'><img src='/images/pflogo.gif' border=0></a></td>"
'The users First and Last Name is Displayed in a hello message.

	Response.Write "<td VAlign=top>"
	%>
	<!--#include file="DateTime.Asp"-->
	
	<%

	IF Session("WorkID")="CB" or Session("WorkID")="JL" then
		Set objConn=Server.CreateObject("ADODB.Connection")
		objConn.Open "DSN=HelpDesk2"
		EndMonth = Month(Date)
		EndDay =Day(Date)
		EndYear = Year(Date)		
		StrDate = EndMonth & "/" & EndDay & "/" & EndYear
		Set rsCount=objConn.Execute("Select Count(*) from Calls WHERE Closed = 'No' and CALL_SERVICED_BY = '"& Session("WorkID") &"'")
		Set rsClosedCount=objConn.Execute("Select Count(*) from Calls WHERE Closed = 'Yes' and Date_Closed = #" & StrDate & "# and CALL_SERVICED_BY = '"& Session("WorkID") &"'")
		Response.Write "<p>Hello " & objRs("FirstName") &" " & objRs("LastName") &"<p>Total Tickets Open: " & rsCount(0) & "<p>Total Tickets Closed Today: " & rsClosedCount(0) &"</font>"
		Response.Write "</td></tr><tr>"
	Else
		Response.Write "<p>Hello " & objRs("FirstName") &" " & objRs("LastName") &"</font>"
		Response.Write "</td></tr><tr>"
	End IF
'The User ID is used as a variable that is passed to the Password Change Page.
	Response.Write "<td valign=top width=125 bgcolor=white>"
	%>
	<!--#include file="_Nav.Asp"-->
	<%
Response.Write "</td>"
	%>
