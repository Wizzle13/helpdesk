<font face="Arial" size="3">
<%
Dim strSurvey_Date
Dim strSurvey_Time
Dim strTicket_Number
Dim strTimely_Manner
Dim strTimely_Manner_Expectation
Dim strProblem_Resolution
Dim strOverall_Experience
Dim strTreated_Answer
Dim strTreated_Answer_No
Dim strOther_Comments
Dim objCDOMail
Dim objConn
Dim objRec
Dim sql
Dim objIdentity
Dim ObjRs
Dim ObjRs2
Dim ObjRs3

If Request("Num") = "" Then
	TickNum = 0
Else
	TickNum = Request("Num")
End If

strSurvey_Date = Request.Form("Survey_Date")
strSurvey_Time = Request.Form("Survey_Time")
strTicket_Number = Request.Form("Ticket_Number")
strTimely_Manner = Request.Form("Timely_Manner")
strTimely_Manner_Expectation = Request.Form("Timely_Manner_Expectation")
strProblem_Resolution = Request.Form("Problem_Resolution")
strOverall_Experience = Request.Form("Overall_Experience")
strTreated_Answer = Request.Form("Treated_Answer")
strTreated_Answer_No = Request.Form("Treated_Answer_No")
strOther_Comments = Request.Form("Other_Comments")

'Response.Write "Date: " & strSurvey_Date
'Response.Write "<BR>Time: " & strSurvey_Time
'Response.Write "<BR>Ticket Number: " & strTicket_Number
'Response.Write "<BR>Timely Manner: " & strTimely_Manner
'Response.Write "<BR>Timely Manner_Expectation: " & strTimely_Manner_Expectation
'Response.Write "<BR>Overall Experience: " & strOverall_Experience
'Response.Write "<BR>Treated Answer: " & strTreated_Answer
'Response.Write "<BR>Treated Answer NO: " & strTreated_Answer_No
'Response.Write "<BR>Other_Comments: " & strOther_Comments

Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
	
objConn.Open "HelpDesk2"
sqlId ="SELECT * FROM Survey WHERE Ticket_Number=" & TickNum & ";"
objRec.Open sqlId, objConn

If objRec.EOF Then
	sql = "INSERT INTO Survey(Ticket_Number, Survey_Date, Survey_Time, Timely_Manner, Timely_Manner_Expectation, Problem_Resolution, Overall_Experience, Treated_Answer, Treated_Answer_No, Other_Comments) VALUES('" & strTicket_Number & "','" & strSurvey_Date & "','" & strSurvey_Time & "','" & strTimely_Manner & "','" & strTimely_Manner_Expectation & "','" & strProblem_Resolution & "','" & strOverall_Experience & "','" & strTreated_Answer & "','" & strTreated_Answer_No & "','" & strOther_Comments & "');"
	objConn.Execute(sql)
	Response.Write "<P>Thanks for filling out this survey for ticket #" & strTicket_Number & "<P>"
	
Else
	sql = "UPDATE Survey SET Survey_Date = '" & strSurvey_Date & "', Survey_Time = '" & strSurvey_Time & "', Timely_Manner = '" & strTimely_Manner & "', Timely_Manner_Expectation = '" & strTimely_Manner_Expectation & "', Problem_Resolution = '" & strProblem_Resolution & "', Overall_Experience = '" & strOverall_Experience & "', Treated_Answer = '" & strTreated_Answer & "', Treated_Answer_No= '" & strTreated_Answer_No & "', Other_Comments= '" & strOther_Comments & "' WHERE Ticket_Number=" & strTicket_Number & ";"
	objConn.Execute(sql)
 	Response.Write "<P>Your survey has been updated for ticket #" & strTicket_Number & "<P>"		
End If

	
objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing

%>

Thank you for your feedback!
