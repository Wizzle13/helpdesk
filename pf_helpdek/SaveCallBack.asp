<%

Dim objConn
Dim objRec
Dim objRs
Dim sql
Dim sql2
Dim SolutionDesc
Dim strDate
Dim DTStamp

Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
objConn.Open "DSN=HelpDesk2"
objRec.Open "Calls", objConn
strCallBack = Request.form("Call_Back_Complete")
strComments = Replace(Request.form("Comments"), "'", "''")
strComments = strComments + vbcrlf  + "---------------------" +  Session("FirstName") + " " + Session("LastName") + " - " + CStr(Now) + vbcrlf + vbcrlf
'Solution = SolutionDesc


If Request.form("Call_Back_Complete") ="on" Then
	sql = "UPDATE Calls SET  Call_Back_Comments= '" & StrComments & "', Call_Back='Yes', Call_Back_Date= '" & CStr(Date) & "', Call_Back_Time= '" & CStr(Time) & "' WHERE Ticket_Number=" & Request("Num") & ";"
Else
	sql = "UPDATE Calls SET  Call_Back_Comments= '" & StrComments & "' WHERE Ticket_Number=" & Request("Num") & ";"
End IF

objConn.Execute(sql)
	
objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing
TicketNum=Request("Num") 
	
Response.Redirect "CallBack.asp"
%>

