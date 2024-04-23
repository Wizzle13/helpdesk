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

sqlId ="SELECT * FROM GROUP_MEMBERS WHERE PACK_ID=" & Request("Num") & ";"

objRec.Open sqlId, objConn

Response.Write "<a href=delete.asp?Num=" & objRec("PACK_ID") & ">Delete</a>"	

Response.Write "<form method=post action=save.asp?Num=" & Request("Num") & " name='Join_Form1'>"

Response.Write "<table border=0>"
Response.Write "<tr><td>Category:</td><td><select name=Category>"
If objRec("HW_SW_ID") = "HW" Then
	Response.Write "<option value=HW selected>Hardware"
	Response.Write "<option value=SW>Software</select><BR>"
Else
	Response.Write "<option value=HW>Hardware"
	Response.Write "<option value=SW selected>Software</select><BR>"
End If
Response.Write "</td></tr>"
Response.Write "<tr><td>Group:</td><td><input name=Group value='" & objRec("GROUP_NAME") & "'></td></tr>"
Response.Write "<tr><td>Package:</td><td><input name=Package value='" & objRec("PACKAGE_NAME") & "'></td></tr>"
Response.Write "<tr><td colspan=2><p><input type=Submit  value='Update Information'></form></td></tr></table>"

%>

</BODY>
</HTML>