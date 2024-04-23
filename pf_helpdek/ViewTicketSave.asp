<%
Dim objConn
Dim objRec
Dim sql
Dim sql2
Dim SolutionDesc
	
Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
objConn.Open "DSN=HelpDesk2"
objRec.Open "Calls", objConn

ProblemDesc = Replace(Request.form("Problem_Desc"), "'", "''")
ProblemDesc2 = Replace(Request.form("Problem_Desc2"), "'", "''")
ProblemDesc = ProblemDesc + vbcrlf  + "---------------------" + vbcrlf + CStr(now) + vbcrlf + "---------------------" + vbcrlf  + ProblemDesc2
SolutionDesc = Replace(Request.form("Solution_Desc"), "'", "''")
SolutionDesc = SolutionDesc + vbcrlf  + "---------------------" +  Session("FirstName") + " " + Session("LastName") + " - " + CStr(now) + vbcrlf + vbcrlf
ROI = Replace(Request.form("ROI"), "'", "''")
Problem = Request.form("Problem_Desc")
Solution = Request.form("Solution_Desc")
strROI = Request.form("ROI")


If Request.Form("Date_Closed") = "" then
	strDate= Request.Form("Date_Closed2")
Else
	strDate= Request.Form("Date_Closed")
end if
If Request.Form("Time_Closed") = "" then
	strTime= Request.Form("Time_Closed2")
Else
	strTime= Request.Form("Time_Closed")
end if
If Request.Form("Closed") = "Yes" then
	SolutionDesc= SolutionDesc +  vbcrlf + "Closed by user."
end if

sql = "UPDATE Calls SET  Problem_Desc= '" & ProblemDesc & "', Solution_Desc= '" & SolutionDesc & "', Closed= '" & Request.Form("Closed") & "', Date_Closed= '" & strDate & "', Time_Closed= '" & strTime & "', ROI= '" & ROI & "' WHERE Ticket_Number=" & Request("Num") & ";"

objConn.Execute(sql)
	
objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing
TicketNum=Request("Num") 




	Set objConn=Server.CreateObject("ADODB.Connection")
	objConn.Open "DSN=Helpdesk2"
	Set objRs=objConn.Execute("Select * from UserInfo WHERE WorkerID = '"& Request.Form("Call_Serviced_By") &"'")
	Set objRs2=objConn.Execute("Select * from Domains WHERE Country = '"& objRs("Country") &"'")
	strSend =objRs("Email") & objRs2("Domain")
	If Request.Form("Closed") = "Yes" then
		strBody="Help Desk ticket #"+ TicketNum +" has been closed by " + Session("FirstName") & " " & Session("LastName") & "." + vbcrlf + vbcrlf + "You can see this ticket at the following address: http://172.16.3.10/modifyticket.asp?Num=" + TicketNum + vbcrlf
	Else
		strBody="Help Desk ticket #"+ TicketNum +" has been modified by " + Session("FirstName") & " " & Session("LastName") & "." + vbcrlf + vbcrlf + "You can see this ticket at the following address: http://172.16.3.10/modifyticket.asp?Num=" + TicketNum + vbcrlf
	End If
	objRs.close
	objConn.Close
	Set objRs = Nothing
	Set objConn = Nothing
	Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
    
	' Set the properties of the object
	objCDOMail.From = "support@purefishing.com"
	objCDOMail.To = strSend
	If Request.Form("Closed") = "Yes" then
		objCDOMail.Subject = "Ticket " & TicketNum & " has been closed."
	else
		objCDOMail.Subject = "Ticket " & TicketNum & " has been updated."
	End if
	objCDOMail.Body = strBody

	objCDOMail.Send

	Set objCDOMail = Nothing  	


If Session("PageName") = "Search"	Then
	Response.Redirect "Search.asp"
Else	
	Response.Redirect "Main.asp"
End IF	
%>

