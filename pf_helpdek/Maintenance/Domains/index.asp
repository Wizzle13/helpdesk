<!--#include file="../../_validsession.asp"-->
<!--#include file="../../_head.asp"-->

<!--
Programer: Chris Burton    
Date Started: 10-22-02

Description:
This page is the Manintenance Section for the Domains.
-->
<td  valign="top">
<form method="post" action="Add.asp"><input type="submit" value="Add" name="B1"> 
</form>
<%
Set objConn=Server.CreateObject("ADODB.Connection")
	objConn.Open "DSN=Helpdesk2"
	Set objRs=objConn.Execute("Select * from Domains order by CountryFull Asc,Domain")
	sqlId ="SELECT * FROM Domains;"
	Response.Write"<table border='1'><tr><td><b>Country</b></td><td><b>Domain</b></td></tr>"
	While Not objRs.EOF
	
	Response.Write"<tr><td><a href='modify.asp?num=" & objRs("ID") & "'>"& objRs("CountryFull")&"</a></td><td>"&objRs("Domain") &"</td></tr> "
		objRs.MoveNext
	Wend

%>
</td>
</tr>
</table>
</body>
</html>