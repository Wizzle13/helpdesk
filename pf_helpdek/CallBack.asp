<!--#include file="_head.asp"-->
<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"

StartMonth = 1
StartDay = 1
StartYear = 2006
Start_Date = StartMonth & "/" & StartDay & "/" & StartYear

Call_Back_Date = Date 
Call_Back_Date1 = Date - 1
Call_Back_Date2 = Date - 2
Call_Back_Date3 = Date - 14		


Set objRs=objConn.Execute("Select * from Calls WHERE Closed = 'Yes' and Call_Type ='Help' and CALL_SERVICED_BY =  '"& Session("WorkID") &"' and (Date_Closed >= #" & Start_Date & "#) and (Date_Closed <= #" & Call_Back_Date & "#) and Call_Back = 'No'  order by Date_Closed, Time_Closed")
Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Closed = 'No' and CALL_SERVICED_BY =  '"& Session("WorkID") &"'")
Response.Write "<td>"

If rsCount(0) <> 0 Then
Response.Write "<table border=0><TR>"
Response.Write "<TD bgcolor=#FBFD04>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>< 2 Days (" & Call_Back_Date1 & ")</TD><TD>&nbsp;&nbsp;</TD>"
Response.Write "<TD bgcolor=#A8EAA8>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>>= 2 Days(" & Call_Back_Date2 & ")</TD><TD>&nbsp;&nbsp;</TD>"
Response.Write "<TD bgcolor=#FF6666>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>> 14 Days(" & Call_Back_Date3 & ")</TD></TR></TABLE>"

Response.Write "<table border=1 cellpadding=0 cellborder=0>"
Response.Write "<tr><td><B>Ticket #</td>"
Response.Write "<td><B>User<br>Name</td>"
Response.Write "<td><B>Contact #</td>"
Response.Write "<td><B>Date<BR>Closed</td>"
'Response.Write "<td><B>Serviced<br>By</td>"
'Response.Write "<td width= 50><B>Problem</td>"


	While Not objRs.EOF
		Response.Write "<tr><td align=center><a href=""ModifyCallBack.asp?Num=" & objRs("Ticket_Number") & """>" & objRs("Ticket_Number") & "</a>" & "</td>"
		Response.Write "<td>" & objRs("User_First_Name") &"<br>"& objRs("User_Last_Name") & "</td>"	
		IF objRs("Date_Closed") => Call_Back_Date1 Then
			Response.Write "<td bgcolor=#FBFD04 align=center><font color=black>"& objRs("USER_CONTACT_NUMBER") & "</td>"
		End IF	
		IF objRs("Date_Closed") <= Call_Back_Date2  and objRs("Date_Closed") > Call_Back_Date3 Then
			Response.Write "<td bgcolor=#A8EAA8 align=center><font color=black>"& objRs("USER_CONTACT_NUMBER") & "</td>"
		End IF	
		If objRs("Date_Closed") <= Call_Back_Date3 Then
			Response.Write "<td bgcolor=#FF6666 align=center><font color=black>"& objRs("USER_CONTACT_NUMBER") & "</td>"
		End IF	
		Response.Write "<td>" & objRs("Date_Closed") & "<br>"& objRs("Time_Closed") &"</td>"
		'Response.Write "<td>" & objRs("CALL_SERVICED_BY") & "</td>"						

'		ProblemDesc = objRs("PROBLEM_DESC")
'		If Len(ProblemDesc) > 100 Then
'			Response.Write "<td>" & Mid(ProblemDesc, 1, 100) & "...<a href=""ModifyCalBack.asp?Num=" & objRs("Ticket_Number") & """>(more)</a></td>"	
'		Else
'			Response.Write "<td>" & ProblemDesc & "</td>"
'		End If		

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
set objConn=nothing

Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

%>
</body>
</html>