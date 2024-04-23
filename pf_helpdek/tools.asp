<!--#include file="_head.asp"-->
<%
	Response.Write "<td>"
	Response.Write "<a href=tickets.asp?Num=1>Open Help Tickets</a><p>"
	Response.Write "<a href=tickets.asp?Num=2>Open Request Tickets</a><p>"
	Response.Write "<a href=tickets.asp?Num=3>Closed Help Tickets</a><p>"
	Response.Write "<a href=tickets.asp?Num=4>Close Request Tickets</a><p>"
	Response.Write "<a href=tickets.asp?Num=5>All Help Tickets</a><p>"
	Response.Write "<a href=tickets.asp?Num=6>All Request Tickets</a><p>"
	Response.Write "<a href=tickets.asp?Num=7>All Open Tickets</a><p>"
	Response.Write "<a href=tickets.asp?Num=8>All Closed Tickets</a><p>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

%>
</body>
</html>