<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"
Set objRs=objConn.Execute("Select * from Calls WHERE Closed = 'No' and Email =  '"& Session("Email")&"' and Country='" & Session("Country") &"' order by Ticket_Number")
Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Closed = 'No' and Email =  '"& Session("Email")&"' and Country='" & Session("Country") &"'")

Response.Write "<B>MY TICKETS</B>"

If rsCount(0) <> 0 Then
	Response.Write "<table border=0><TR><TD bgcolor=#A8EAA8>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>In Process</TD><TD>&nbsp;&nbsp;</TD>"
	Response.Write "<TD bgcolor=#FBFD04>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Pending</TD><TD>&nbsp;&nbsp;</TD>"
	Response.Write "<TD bgcolor=#FF6666>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Stopped</TD><TD>&nbsp;&nbsp;</TD>"
	Response.Write "<TD bgcolor=lightgrey>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Closed</TD></TR></TABLE>"

	Response.Write "<table border=1 cellpadding=0 cellborder=0>"
	Response.Write "<tr><td><B>Ticket #</td>"
	Response.Write "<td><B>User<br>Name</td>"
	Response.Write "<td><B>Call<br>Type</td>"
	Response.Write "<td><B>Date<BR>Opened</td>"
	Response.Write "<td><B>Serviced<br>By</td>"
	Response.Write "<td width= 50><B>Problem</td>"

	While Not objRs.EOF
		Response.Write "<tr><td align=center><a href=""Viewticket.asp?Num=" & objRs("Ticket_Number") & """>" & objRs("Ticket_Number") & "</a>" & "</td>"
		Response.Write "<td>" & objRs("User_First_Name") &"<br>"& objRs("User_Last_Name") & "</td>"	
		If objRs("IN_PROCESS") = "Yes" Then
			Response.Write "<td bgcolor=#A8EAA8 align=center><font color=black>"& objRs("Call_Type") & "</td>"
		End If
		If objRs("IN_PROCESS") = "MayBe" Then
			Response.Write "<td bgcolor=#FBFD04 align=center><font color=black>"& objRs("Call_Type") & "</td>"
		End If
		If objRs("IN_PROCESS") = "No" Then
			Response.Write "<td bgcolor=#FF6666 align=center><font color=black>"& objRs("Call_Type") & "</td>"
		End If
		Response.Write "<td>" & objRs("Date_Opened") & "<br>"& objRs("Time_Opened") &"</td>"
		Response.Write "<td>" & objRs("CALL_SERVICED_BY") & "</td>"						
		
		ProblemDesc = objRs("PROBLEM_DESC")
		If Len(ProblemDesc) > 100 Then
			Response.Write "<td>" & Mid(ProblemDesc, 1, 100) & "...<a href=""viewticket.asp?Num=" & objRs("Ticket_Number") & """>(more)</a></td></tr>"	
		Else
			Response.Write "<td>" & ProblemDesc & "</td></tr>"
		End If	
	
		objRs.MoveNext
	Wend
Else
	Response.Write "<table border=0><TR><TD bgcolor=#A8EAA8>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>In Process</TD><TD>&nbsp;&nbsp;</TD>"
	Response.Write "<TD bgcolor=#FBFD04>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Pending</TD><TD>&nbsp;&nbsp;</TD>"
	Response.Write "<TD bgcolor=#FF6666>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Stopped</TD><TD>&nbsp;&nbsp;</TD>"
	Response.Write "<TD bgcolor=lightgrey>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Closed</TD></TR></TABLE>"

	Response.Write "<table border=1 cellpadding=0 cellborder=0>"
	Response.Write "<tr><td><B>Ticket #</td>"
	Response.Write "<td><B>User<br>Name</td>"
	Response.Write "<td><B>Call<br>Type</td>"
	Response.Write "<td><B>Date<BR>Opened</td>"
	Response.Write "<td><B>Serviced<br>By</td>"
	Response.Write "<td width= 50><B>Problem</td>"
End If

objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing

'**************** Display last 5 tickets. ***********

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"
Set objRs=objConn.Execute("Select TOP 5 * from Calls WHERE Closed = 'Yes' and Email =  '"& Session("Email")&"' and Country='" & Session("Country") &"' order by Ticket_Number desc")

While Not objRs.EOF
	Response.Write "<tr><td align=center bgcolor=lightgrey><a href=""Viewticket.asp?Num=" & objRs("Ticket_Number") & """>" & objRs("Ticket_Number") & "</a>" & "</td>"
	Response.Write "<td bgcolor=lightgrey>" & objRs("User_First_Name") &"<br>"& objRs("User_Last_Name") & "</td>"	
	If objRs("IN_PROCESS") = "Yes" Then
		Response.Write "<td bgcolor=#A8EAA8 align=center><font color=black>"& objRs("Call_Type") & "</td>"
	End If
	If objRs("IN_PROCESS") = "MayBe" Then
		Response.Write "<td bgcolor=#FBFD04 align=center><font color=black>"& objRs("Call_Type") & "</td>"
	End If
	If objRs("IN_PROCESS") = "No" Then
		Response.Write "<td bgcolor=#FF6666 align=center><font color=black>"& objRs("Call_Type") & "</td>"
	End If
	Response.Write "<td bgcolor=lightgrey>" & objRs("Date_Opened") & "<br>"& objRs("Time_Opened") &"</td>"
	Response.Write "<td bgcolor=lightgrey>" & objRs("CALL_SERVICED_BY") & "</td>"
						
	ProblemDesc = objRs("PROBLEM_DESC")
	If Len(ProblemDesc) > 100 Then
		Response.Write "<td bgcolor=lightgrey>" & Mid(ProblemDesc, 1, 100) & "...<a href=""viewticket.asp?Num=" & objRs("Ticket_Number") & """>(more)</a></td></tr>"	
	Else
		Response.Write "<td bgcolor=lightgrey>" & ProblemDesc & "</td></tr>"
	End If	
				
	objRs.MoveNext
Wend

Response.Write "</table><br>"
Response.Write "<a href=ViewAllMyTickets.asp>View All My Tickets</a><P>"
Response.Write "<a href='main.asp#PageTop'>Top</a><p>"

objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing
%>

