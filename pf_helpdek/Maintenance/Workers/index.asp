<!--#include file="../../_validsession.asp"-->
<!--#include file="../../_head.asp"-->

<!--
Programer: Chris Burton    
Date Started: 3-14-02

Description:
This page is the Manintenance Section for the workers.
-->
<td>
<%
Set objConn=Server.CreateObject("ADODB.Connection")
	objConn.Open "DSN=HelpDesk2"
	Set objRs=objConn.Execute("Select * from Workers  order by Active_worker Desc,First_Name, Last_Name")
	sqlId ="SELECT * FROM Workers;"
	Response.Write"<table border='1'><tr><td><b>First Name</b></td><td><b>Last Name</b></td><td><b>Active Worker?</b></td></tr>"
	While Not objRs.EOF
	
	Response.Write"<tr><td>"& objRs("First_Name")&"</td><td>"& objRs("Last_Name") &"</td><td>"&objRs("Active_Worker") &"</td></tr> "
		objRs.MoveNext
	Wend

%>
</td>
</tr>
</table>
</body>
</html>