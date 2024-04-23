<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"

Set objRs=objConn.Execute("Select * from Calls WHERE Closed = 'No' and CALL_SERVICED_BY =  '"& Session("WorkID") &"' order by Priority, Projected_Complete_Date, Ticket_Number")
Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Closed = 'No' and CALL_SERVICED_BY =  '"& Session("WorkID") &"'")

Response.Write "<B>ASSIGNED TICKETS</B> <a href='printmytickets.asp'>Printable List</a>"
If rsCount(0) <> 0 Then
Response.Write "<table border=0><TR><TD bgcolor=#A8EAA8>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>In Process</TD><TD>&nbsp;&nbsp;</TD>"
Response.Write "<TD bgcolor=#FBFD04>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Pending</TD><TD>&nbsp;&nbsp;</TD>"
Response.Write "<TD bgcolor=#FF6666>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Stopped</TD></TR></TABLE>"
Response.Write "<table border=1 cellpadding=0 cellborder=0>"
Response.Write "<tr><td><B>Ticket #</td>"
Response.Write "<td><B>SAP<br>Priority</td>"
Response.Write "<td><B>User<br>Name</td>"
Response.Write "<td><B>Call<br>Type</td>"
Response.Write "<td><B>Date<BR>Opened</td>"
Response.Write "<td><B>Projected<br>Complete<br>Date</td>"
'Response.Write "<td><B>Serviced<br>By</td>"
Response.Write "<td width= 50><B>Problem</td>"
		
	While Not objRs.EOF
		Response.Write "<tr><td align=center><a href=""modifyticket.asp?Num=" & objRs("Ticket_Number") & """>" & objRs("Ticket_Number") & "</a>" & "</td>"
		Response.Write "<td align=center>" & objRs("SAPPriority") & "</td>"
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
		IF objRs("Projected_Complete_Date") = "12/31/2049" then
			Response.Write "<td></td>"
		Else
			Response.Write "<td>" & objRs("Projected_Complete_Date") &"</td>"
		End IF	
		'Response.Write "<td>" & objRs("CALL_SERVICED_BY") & "</td>"						

		ProblemDesc = objRs("PROBLEM_DESC")
		If Len(ProblemDesc) > 100 Then
			Response.Write "<td>" & Mid(ProblemDesc, 1, 100) & "...<a href=""modifyticket.asp?Num=" & objRs("Ticket_Number") & """>(more)</a></td>"	
		Else
			Response.Write "<td>" & ProblemDesc & "</td>"
		End If		

		'************************************
		'*  Fill in the Development Hours column, even if there isn't one in the database
		'************************************
		
		objRs.MoveNext
	Wend

	Response.Write "</table><P>"
Else
	Response.Write "<BR>You are not currently working on any tickets.<P>"
End If
Response.Write "<a href='main.asp#PageTop'>Top</a><p>"
objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing%>