<!--#include file="_head.asp"-->

<script language="JavaScript">
function go_worker(workerdd) {
	location.href = 'saptickets.asp?worker=' + workerdd.value;
}
function save_ticket() {
	
}

</script>
<%
Dim ServiceID
 
ServiceID = Request("worker")

If ServiceID = "" Then
	ServiceID = Session("WorkID")
End if


Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"

'Response.Write Request("Worker")
Pack_Name = "SAP"
If Request("Worker") ="All" Then
	Set objRs=objConn.Execute("Select * from Calls WHERE Closed = 'No' AND Package_Name Like '%" & Pack_Name &"%' order by TransportDate, SAPPriority, Ticket_Number")
	Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Closed = 'No' AND Package_Name Like '%" & Pack_Name &"%'")
Else
	Set objRs=objConn.Execute("Select * from Calls WHERE Closed = 'No' AND Package_Name Like '%" & Pack_Name &"%' and CALL_SERVICED_BY =  '" & ServiceID & "' order by TransportDate, SAPPriority, Ticket_Number")
	Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Closed = 'No' AND Package_Name Like '%" & Pack_Name &"%' and CALL_SERVICED_BY = '" & ServiceID & "'")
End If


Response.Write "<form name=workerdd method=post>"  & vbcrlf

Set objConn2=Server.CreateObject("ADODB.Connection")
objConn2.Open "DSN=HelpDesk2"

If ServiceID <> "All" Then
	Set objRec=objConn.Execute("Select * from Workers where Worker_ID='" & ServiceID & "'")
	strWorker= objRec("Worker_ID")

	Set objRs2=objConn2.Execute("Select * from Workers where Worker_ID= '"&strWorker&"'")
End if
Response.Write "<td valign=top><font size=1 face=Arial><p>Person Working Call:" & vbcrlf
Response.Write "<font size=1 face=Arial><select name=Worker_ID onchange=go_worker(this)>" & vbcrlf

Response.Write "<Option value=All selected>All Open SAP Issues"

Set objRs2=objConn2.Execute("Select * from Workers where Active_Worker='Yes' order by First_Name, Last_Name")
sqlId ="SELECT * FROM Workers;"
While Not objRs2.EOF
	If strWorker = objRs2("Worker_ID") Then
		Response.Write"<Option value=" & objRs2("Worker_ID") & " selected>"& objRs2("First_Name")&" "& objRs2("Last_Name") & vbcrlf
	Else
		Response.Write"<Option value="& objRs2("Worker_ID")&">"& objRs2("First_Name")&" "& objRs2("Last_Name") & vbcrlf
	End if
	objRs2.MoveNext
Wend
objRs2.Close
Set objRs2 = Nothing
objConn2.Close
Set objConn2 = Nothing

Response.Write"</select></form>" & vbcrlf
Response.Write "You are viewing " & RsCount(0) & " tickets."

Response.Write "<a href='printsaptickets.asp?worker="& Request("worker") & "'>Printable List</a>" & vbcrlf
If rsCount(0) <> 0 Then
Response.Write "<table border=0><TR><TD bgcolor=#A8EAA8>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>In Process</TD><TD>&nbsp;&nbsp;</TD>" & vbcrlf
Response.Write "<TD bgcolor=#FBFD04>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Pending</TD><TD>&nbsp;&nbsp;</TD>" & vbcrlf
Response.Write "<TD bgcolor=#FF6666>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Stopped</TD></TR></TABLE>" & vbcrlf
Response.Write "<table border=1 cellpadding=0 cellborder=0>"
Response.Write "<tr><td><B>Ticket #</td>" & vbcrlf
Response.Write "<td><B>Transport Date</td>" & vbcrlf
Response.Write "<td><B>Priority</td>" & vbcrlf
Response.Write "<td><B>User<br>Name</td>" & vbcrlf
Response.Write "<td><B>Call<br>Type</td>" & vbcrlf
Response.Write "<td><B>Date<BR>Opened</td>" & vbcrlf
If Request("Worker") ="All" Then
	Response.Write "<td><B>Serviced<br>By</td>" & vbcrlf
End if
Response.Write "<td width= 50><B>Problem</td>" & vbcrlf

f = 1 
		
	While Not objRs.EOF
		Response.Write "<tr><td align=center><a href=""modifyticket.asp?Num=" & objRs("Ticket_Number") & """>" & objRs("Ticket_Number") & "</a>" & "</td>"
		IF objRs("TransportDate") ="12/31/2049" or objRs("TransportDate") ="" Then
			Response.Write "<td align=center>N/A</td>" & vbcrlf
		Else
			Response.Write "<td align=center>" & objRs("TransportDate") & "</td>" & vbcrlf
		End IF		
		
		Response.Write "<td align=center><form name=frmpriority" & f & " method=post action=savesapticket.asp?Num=" & objRs("Ticket_Number") & ">" & vbcrlf
		Response.Write "<script language=JavaScript>"
		Response.Write "function save_ticket" & f & "(){"
		Response.Write "document.frmpriority" & f & ".submit();"
		Response.Write "}"
		Response.Write "</script>"
		
		Response.Write "<select name=tickpriority onchange=save_ticket" & f & "()>" & vbcrlf
		f = f + 1
		i = 1
		While i <> 21
			If i = objRs("SAPPriority") Then
				Response.Write "<option value=" & objRs("SAPPriority") & " selected>" & objRs("SAPPriority") & vbcrlf
			Else
				Response.Write "<option value=" & i & ">" & i & vbcrlf
			End if
			i = i + 1
		Wend
		If objRs("SAPPriority") = 99 Then
			Response.Write "<option value=" & objRs("SAPPriority") & " selected>" & objRs("SAPPriority") & vbcrlf
		Else
			Response.Write "<option value=99>99" & vbcrlf
		End if
		Response.Write "</select></form></td>"  & vbcrlf
		Response.Write "<td>" & objRs("User_First_Name") &"<br>"& objRs("User_Last_Name") & "</td>" & vbcrlf
		If objRs("IN_PROCESS") = "Yes" Then
					Response.Write "<td bgcolor=#A8EAA8 align=center><font color=black>"& objRs("Call_Type") & "</td>" & vbcrlf
		End If
		If objRs("IN_PROCESS") = "MayBe" Then
					Response.Write "<td bgcolor=#FBFD04 align=center><font color=black>"& objRs("Call_Type") & "</td>" & vbcrlf
		End If
		If objRs("IN_PROCESS") = "No" Then
					Response.Write "<td bgcolor=#FF6666 align=center><font color=black>"& objRs("Call_Type") & "</td>" & vbcrlf
		End If
		Response.Write "<td>" & objRs("Date_Opened") & "<br>"& objRs("Time_Opened") &"</td>" & vbcrlf
		If Request("Worker") ="All" Then
			Response.Write "<td>" & objRs("CALL_SERVICED_BY") & "</td>" & vbcrlf						
		End if
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

<%
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

%>

</body>
</html>