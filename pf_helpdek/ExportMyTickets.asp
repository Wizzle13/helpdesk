<%
Response.Buffer = True
Response.ContentType = "application/vnd.ms-excel"

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"

Set objRs=objConn.Execute("Select * from Calls WHERE Closed = 'No' and CALL_SERVICED_BY =  '"& Session("WorkID") &"' order by Priority, Ticket_Number")

Response.Write "<table border=1 cellpadding=0 cellborder=0>"
Response.Write "<tr><td><B>Ticket #</td>"
Response.Write "<td><B>User Name</td>"
Response.Write "<td><B>Call Type</td>"
Response.Write "<td><B>Date/Time<BR>Opened</td>"
'Response.Write "<td><B>Serviced By</td>"
Response.Write "<td width= 490><B>Problem/Solution</td></tr>"
		
	While Not objRs.EOF
		Response.Write "<tr><td align=center><a href=""modifyticket.asp?Num=" & objRs("Ticket_Number") & """>" & objRs("Ticket_Number") & "</a>" & "</td>"
		Response.Write "<td>" & objRs("User_First_Name") &" "& objRs("User_Last_Name") & "</td>"	
		If objRs("IN_PROCESS") = "Yes" Then
					Response.Write "<td bgcolor=#A8EAA8 align=center><font color=black>"& objRs("Call_Type") & "</td>"
		End If
		If objRs("IN_PROCESS") = "MayBe" Then
					Response.Write "<td bgcolor=#FBFD04 align=center><font color=black>"& objRs("Call_Type") & "</td>"
		End If
		If objRs("IN_PROCESS") = "No" Then
					Response.Write "<td bgcolor=#FF6666 align=center><font color=black>"& objRs("Call_Type") & "</td>"
		End If
		Response.Write "<td>" & objRs("Date_Opened") & " - "& objRs("Time_Opened") &"</td>"
		'Response.Write "<td>" & objRs("CALL_SERVICED_BY") & "</td>"						
		Response.Write "<td><B>Problem: </B>" & objRs("PROBLEM_DESC") & vbcrlf & "<P><B>Solution: </B>" & objRs("SOLUTION_DESC") & "</td></tr>"		
		
		objRs.MoveNext
	Wend

	Response.Write "</table><P>"
objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing%>