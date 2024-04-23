<HTML>
<HEAD>
<TITLE>Pure Fishing Online Help Desk FAQ (Frequently Asked Questions)</TITLE>
</HEAD>
<BODY>
<a name=top>
<a href="/"><img src="/images/pflogo.gif" border=0></a>
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
<a href="FAQ.asp">Back to FAQs</a>
</BODY>
</HTML>
