<!--#include file="../_head.asp"-->

<!--
Programer: Chris Burton    
Date Started: 4-22-03

Description:
This Page is for the New Hire Forms

-->

<body>

<%
Session("PageName") = "New Hire"

	Response.Write "<td valign=top>"
%>	

<font size=3>Please use the appropriate link below:</font><P>
<UL>
<LI><a href="addindex.asp">New/Update Employee Form</A></LI><P>
<LI><a href="termindex.asp">Employee Termination Form</A></LI>
</UL>

<%
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

%>
</body>
</html>

