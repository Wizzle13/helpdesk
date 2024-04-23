<!--#include file="_head.asp"-->
<%

Dim sqlId
Dim objConn
Dim objRec

Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")

objConn.Open "DSN=Helpdesk2"
objRec.Open "WhosOn", objConn	
Set objRec=objConn.Execute("Select * from WhosOn")

	sql = "UPDATE Whoson SET Status= 'Off' WHERE LogonDate <= #" & FormatDateTime (Date -2, vbShortDate) & "#;"
	objConn.Execute(sql)
	
'Response.Redirect "WhosOn.asp"

objRec.Close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
If Request("sortby")="FName" then
Set objRs=objConn.Execute("Select * from WhosOn where Status='On' order by FirstName, LastName")
End if

If Request("sortby")="LName" then
Set objRs=objConn.Execute("Select * from WhosOn where Status='On' order by LastName, FirstName")
End if

If Request("sortby")="Date" or Request("sortby")="" then
Set objRs=objConn.Execute("Select * from WhosOn where Status='On' order by LogonDate desc, LogonTime desc")
End if
Response.Write "<td valign=top>"



Response.Write "<table border=1 cellpadding=0 cellborder=0>"
Response.Write "<td><B><a href=whoson.asp?sortby=FName>First&nbsp;Name</a></td>"
Response.Write "<td><B><a href=whoson.asp?sortby=LName>Last&nbsp;Name</a></td>"		
Response.Write "<td><B><a href=whoson.asp?sortby=Date>Date/Time</a></td>"		
	While Not objRs.EOF

		Response.Write "<tr>"
		Response.Write "<td>" & objRs("FirstName") &"</td>"	
		Response.Write "<td>"& objRs("LastName") & "</td>"
		Response.Write "<td>"& objRs("LogonDate") & " "& objRs("LogonTime") & "</td>"
		objRs.MoveNext
	Wend

	Response.Write "</table><P>" 
	

	strDate=FormatDateTime (Date -2, vbShortDate)
	'Response.Write "<form method=post action=Clear.asp>"
	'Response.Write "<input type=Hidden name=strDate Value=" & strDate &">"
	'Response.Write "<input type=submit value=Clear name=B1>" 
	'Response.Write "</form>"

objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

%>
</body>
</html>