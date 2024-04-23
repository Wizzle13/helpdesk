<!--#include file="../../_validsession.asp"-->
<!--#include file="../../_head.asp"-->

<!--
Programer: Chris Burton    
Date Started: 7-1-02

Description:
This page is the Manintenance Section for the FAQs.
-->
<td valign="top">
<FONT SIZE=4><P><B>Pure Fishing Online Help Desk FAQ (Frequently Asked Questions)</B></P></font>
<form method="post" action="Add.asp"><input type="submit" value="Add" name="B1"> 
</form>
<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"

Set objRs=objConn.Execute("Select * from FAQs order by FAQID")
Response.Write "<ul>"
While Not objRs.EOF
	Response.Write "<li><a href=Modify.asp?num=" & objRs("FAQID") & ">"& objRs("Question") &"</a>"
	objRs.MoveNext
Wend
Response.Write "</ul>"
%>
</BODY>
</HTML>
