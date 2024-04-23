<!--#include file="_head.asp"-->
<%

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
If Request("sortby")="FName" then
Set objRs=objConn.Execute("Select * from WhosOn order by FirstName, LastName")
End if

If Request("sortby")="LName" then
Set objRs=objConn.Execute("Select * from WhosOn order by LastName, FirstName")
End if

If Request("sortby")="Date" or Request("sortby")="" then
Set objRs=objConn.Execute("Select * from WhosOn order by LogonDate desc, LogonTime desc")
End if
Response.Write "<td valign=top>"



Response.Write "<table border=1 cellpadding=0 cellborder=0>"
Response.Write "<td><B><a href=laston.asp?sortby=FName>First&nbsp;Name</a></td>"
Response.Write "<td><B><a href=laston.asp?sortby=LName>Last&nbsp;Name</a></td>"		
Response.Write "<td><B><a href=laston.asp?sortby=Date>Date/Time</a></td>"		
	While Not objRs.EOF

		Response.Write "<tr>"
		Response.Write "<td>" & objRs("FirstName") &"</td>"	
		Response.Write "<td>"& objRs("LastName") & "</td>"
		Response.Write "<td>"& objRs("LogonDate") & " "& objRs("LogonTime") & "</td>"
		objRs.MoveNext
	Wend

	Response.Write "</table><P>" 
	


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