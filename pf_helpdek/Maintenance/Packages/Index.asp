<!--#include file="../../_validsession.asp"-->
<!--#include file="../../_head.asp"-->

<!--
Programer: Chris Burton    
Date Started: 7-1-02

Description:
This page is the Manintenance Section for the FAQs.
-->
<td valign="top">
<FONT SIZE=4><P><B>Pure Fishing Online Help Desk Package Editor</B></P></font>
<form method="post" action="Add.asp"><input type="submit" value="Add" name="B1"> 
</form>
<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"

Set objRs=objConn.Execute("Select * from GROUP_MEMBERS ORDER by PACKAGE_NAME")
Response.Write "<table border=1><TR><td><b>Package Name</b></td><td><b>Group Name</b></td><td><b>Hardware-Software ID</b></td></tr>"
Response.Write ""
While Not objRs.EOF
	Response.Write "<tr><td><a href=Modify.asp?num=" & objRs("PACK_ID") & ">"& objRs("PACKAGE_NAME") &"</a></td><td>" & objRs("GROUP_NAME") & "</td><td>" & objRs("HW_SW_ID") & "</td></tr>"
	objRs.MoveNext
Wend
Response.Write "</table>"
%>
</BODY>
</HTML>
