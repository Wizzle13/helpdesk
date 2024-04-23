<!--#include file="_head.asp"-->

<!--
Programer: Chris Burton    
Date Started: 7-9-01

Description:
This page checks the  database for users First and Last name
then uses it to display a hello message to the user.

-->

<script language="JavaScript">
<!--
var time = null
function move ()  {
window.location = "Main.asp"
}
//-->
</script>
<!--<body onload="timer=setTimeout ('move()',300000)">-->
<body>
<a Name="#PageTop"></a>
<%
Session("PageName") = "Main"
'This section will pull and display the Users Open tickets.
	Response.Write "<td>"
	If Session("status")="Manager" or Session("status")="Worker" or Session("status")="Administrator" Then
	%>
	<!--#include file="Unassigned.asp"-->
	<!--#include file="MyTickets.Asp"-->
	<!--#include file="HelpTickets.Asp"-->
	<%
	Else
%>
	<!--#include file="HelpTickets.Asp"-->
<%
	end if
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

%>
</body>
</html>
