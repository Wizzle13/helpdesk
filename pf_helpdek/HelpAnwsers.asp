<!--#include file="_head.asp"-->
<td valign="top">
<a name=top>
<FONT SIZE=4><P><B>Pure Fishing Online Help Desk FAQ (Frequently Asked Questions)</B></P></font>
<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"

Set objRs=objConn.Execute("Select * from FAQs Where FAQID =" & Request("Num") & ";")
Response.Write "<b>"& objRs("Question") & "</b>"
Response.Write "<ul>"
While Not objRs.EOF
	Response.Write "<li>"& objRs("Anwser")
	objRs.MoveNext
Wend
Response.Write "</ul>"
%>
<a href="Help.asp">Back to FAQs</a>
</td></tr></table>
</BODY>
</HTML>