<!--#include file="../../_validsession.asp"-->
<!--#include file="../../_head.asp"-->

<!--
Programer: Chris Burton    
Date Started: 2-2503

Description:
This page is the Add Page for the Email Domains.
-->
<%
Set objConn=Server.CreateObject("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
objConn.Open "DSN=HelpDesk2"

sqlId ="SELECT * FROM Domains WHERE ID=" & Request("Num") & ";"

objRec.Open sqlId, objConn
Response.Write "<td valign='top'>"
Response.Write "<form method='post' action='save.asp' name='Inputform'>"
Response.Write "<input type='Hidden' Name='Num' Value='" & objRec("ID") & "'>"
Response.Write "Domain:<br><input type='text' name='Domain' size='20' style='FONT-FAMILY: Arial, Geneva, Helvetica, Helv; FONT-SIZE: xx-small; FONT-STYLE: normal' Value='" & objRec("Domain") & "'><p>"
Response.Write "Country Abvr:<br><input type='text' name='Country' size='20' style='FONT-FAMILY: Arial, Geneva, Helvetica, Helv; FONT-SIZE: xx-small; FONT-STYLE: normal' Value='" & objRec("Country") & "'><p>"
Response.Write "Country Full Name:<br><input type='text' name='CountryFull' size='20' style='FONT-FAMILY: Arial, Geneva, Helvetica, Helv; FONT-SIZE: xx-small; FONT-STYLE: normal' Value='" & objRec("CountryFull") & "'><p>"
Response.Write "Help Desk:<br><input type='text' name='HelpDesk' size='20' style='FONT-FAMILY: Arial, Geneva, Helvetica, Helv; FONT-SIZE: xx-small; FONT-STYLE: normal' Value='" & objRec("Helpdesk") & "'><p>"
Response.Write "<input type='submit' value='Save' name='B1'>" 
Response.Write "</form>"
%>