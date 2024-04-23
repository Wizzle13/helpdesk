<!--#include file="../../_validsession.asp"-->
<!--#include file="../../_head.asp"-->

<!--
Programer: Chris Burton    
Date Started: 7-2-02

Description:
This page is the Moodify Page for the FAQs.
-->
<td valign="top">
<FONT SIZE=4><P><B>Pure Fishing Online Help Desk FAQ (Frequently Asked Questions)</B></P></font>
<%
Set objConn=Server.CreateObject("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
objConn.Open "DSN=HelpDesk2"

sqlId ="SELECT * FROM FAQs WHERE FAQID=" & Request("Num") & ";"

objRec.Open sqlId, objConn	
	Response.Write "<form method=post action=save.asp?Num=" & Request("Num") & " name='Join_Form1'>"
	Response.Write "Question:<br><input type=text Size=50 name=Question value='" & objRec("Question")&"'><p>"
	Response.Write "Anwser:<br><textarea cols=60 rows=15 name=Anwser  style=FONT-FAMILY: Arial, Geneva, Helvetica, Helv>" & objRec("Anwser") &" </textarea>"
	Response.Write "<p><input type=Submit  value='Update Information'  style='FONT-FAMILY: Arial, Geneva, Helvetica, Helv'></form>"
	%>
</BODY>
</HTML>