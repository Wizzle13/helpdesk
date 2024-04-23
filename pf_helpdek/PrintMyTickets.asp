<link rel="stylesheet" href="helpdesk.css" type="text/css">
<script>
function excel(){
win2=window.open("exportmytickets.asp","win2")
win2.creator=self
}
</script>
<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"

Set objRs=objConn.Execute("Select * from Calls WHERE Closed = 'No' and CALL_SERVICED_BY =  '"& Session("WorkID") &"' order by Priority, Ticket_Number")
Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Closed = 'No' and CALL_SERVICED_BY =  '"& Session("WorkID") &"'")

Response.Write "<B>ASSIGNED TICKETS</B>"
If rsCount(0) <> 0 Then
Response.Write "<table border=0><TR><TD bgcolor=#A8EAA8>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>In Process</TD><TD>&nbsp;&nbsp;</TD>"
Response.Write "<TD bgcolor=#FBFD04>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Pending</TD><TD>&nbsp;&nbsp;</TD>"
Response.Write "<TD bgcolor=#FF6666>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Stopped</TD><td width=300 align=right>"
Response.Write "<input type=button value='Export to Excel' onClick=excel()></td></TR></TABLE>"
Response.Write "<table border=1 cellpadding=0 cellborder=0>"
Response.Write "<tr><td><B>Ticket #</td>"
Response.Write "<td><B>User<br>Name</td>"
Response.Write "<td><B>Call<br>Type</td>"
Response.Write "<td><B>Date<BR>Opened</td>"
'Response.Write "<td><B>Serviced<br>By</td>"
Response.Write "<td width= 50><B>Problem/Solution</td>"
		
	While Not objRs.EOF
		Response.Write "<tr><td align=center><a href=""modifyticket.asp?Num=" & objRs("Ticket_Number") & """>" & objRs("Ticket_Number") & "</a>" & "</td>"
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
		'Response.Write "<td>" & objRs("CALL_SERVICED_BY") & "</td>"						
		Response.Write "<td><B>Problem: </B>" & objRs("PROBLEM_DESC") & vbcrlf & vbcrlf & "<P><B>Solution: </B>" & objRs("SOLUTION_DESC") & "</td>"		

		'************************************
		'*  Fill in the Development Hours column, even if there isn't one in the database
		'************************************
		
		objRs.MoveNext
	Wend

	Response.Write "</table><P>"
Else
	Response.Write "<BR>You are not currently working on any tickets.<P>"
End If
objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing%>