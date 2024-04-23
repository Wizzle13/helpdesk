<!--#include file="_head.asp"-->
<td valign="Top">
<a name=top>
<FONT SIZE=4><P><B>Pure Fishing Online Help Desk FAQ (Frequently Asked Questions)</B></P></font>
<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"

Set objRs=objConn.Execute("Select * from FAQs order by FAQID")
Response.Write "<ul>"
While Not objRs.EOF
	Response.Write "<li><a href=HelpAnwsers.asp?num=" & objRs("FAQID") & ">"& objRs("Question") &"</a>"
	objRs.MoveNext
Wend
Response.Write "</ul>"
%>
</td></tr></table>
</BODY>
</HTML>
