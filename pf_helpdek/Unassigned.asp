<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"
Country1=session("Country")

Set objRs=objConn.Execute("Select * from Calls WHERE Closed = 'No' and (CALL_SERVICED_BY =null or CALL_SERVICED_BY ='' or CALL_SERVICED_BY ='XX') order by Ticket_Number")
Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Closed = 'No'and (CALL_SERVICED_BY =null or CALL_SERVICED_BY ='') ")

Response.Write "<P>"
Response.Write "<B>UNASSIGNED TICKETS</B>"
If rsCount(0) <> 0 Then
Response.Write "<table border=0><TR><TD bgcolor=#A8EAA8>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>In Process</TD><TD>&nbsp;&nbsp;</TD>"
Response.Write "<TD bgcolor=#FBFD04>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Pending</TD><TD>&nbsp;&nbsp;</TD>"
Response.Write "<TD bgcolor=#FF6666>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Stopped</TD></TR></TABLE>"

Response.Write "<table border=1 cellpadding=0 cellborder=0>"
Response.Write "<tr><td><B>Ticket #</td>"
Response.Write "<td><B>User<br>Name</td>"
Response.Write "<td><B>Call<br>Type</td>"
Response.Write "<td><B>Date<BR>Opened</td>"
Response.Write "<td><B>Serviced<br>By</td>"
Response.Write "<td width= 50><B>Problem</td>"
f = 0		
	While Not objRs.EOF
		f = f + 1
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

Response.Write "<td align=center>" & vbcrlf & vbcrlf
Response.Write "<script language=JavaScript>" & vbcrlf
Response.Write "function save_ticket" & f & "(){" & vbcrlf
Response.Write "document.frmCallServicedBy" & f & ".submit();" & vbcrlf
Response.Write "}" & vbcrlf
Response.Write "</script>" & vbcrlf
Response.Write "<form name=frmCallServicedBy" & f & " method=post action=saveunassignedticket.asp?Num=" & objRs("Ticket_Number") & ">" & vbcrlf

strWorker= objRs("Call_Serviced_By")

Response.Write "<input type=Hidden Name=strworker Value=" & objRs("Call_Serviced_By") &">" & vbcrlf
Response.Write "<input type=Hidden Name=UFName Value=" & objRs("User_First_Name") & ">" & vbcrlf
Response.Write "<input type=Hidden Name=ULName Value=" & objRs("User_Last_Name") & ">" & vbcrlf

'**********************************************************
'*                                                        *
'*    Uncomment the following for the drop down list.     *
'*                                                        *
'**********************************************************

'Set objConn2=Server.CreateObject("ADODB.Connection")
'objConn2.Open "DSN=HelpDesk2"
'Set objRs2=objConn.Execute("Select * from Workers where Worker_ID= '"&strWorker&"'")


'Response.Write"<select name=Call_Serviced_By onchange=save_ticket" & f & "()>" & vbcrlf
'		If strWorker<>"" Then
'			Response.Write"<Option value= " & objRs2("Worker_ID") & ">"& objRs2("First_Name")&" "& objRs2("Last_Name") & vbcrlf
'		Else
'			Response.Write"<Option>" & vbcrlf
'		End if

'		Set objRs2=objConn2.Execute("Select * from Workers where Active_Worker='Yes' order by First_Name, Last_Name")
'		sqlId ="SELECT * FROM Workers;"
'		While Not objRs2.EOF
'			Response.Write"<Option value="& objRs2("Worker_ID")&">"& objRs2("First_Name")&" "& objRs2("Last_Name") & vbcrlf
'			objRs2.MoveNext
'		Wend
'		Response.Write"</select></form></td>" & vbcrlf & vbcrlf
'		objRs2.Close
'		objConn2.Close
'		Set objRs2 = Nothing
'		Set objConn2 = Nothing

'**********************************************************

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
	Response.Write "<BR>There are no unassigned tickets.<P>"
End If
Response.Write "<a href='main.asp#PageTop'>Top</a>"
objRs.close
objConn.Close
Set objRs= Nothing
set objConn=nothing
%>
<P>
