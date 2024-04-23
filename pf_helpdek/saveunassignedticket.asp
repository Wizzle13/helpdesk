<%
Dim objConn
Dim objRec
Dim objRs
Dim sql
Dim sql2
Dim TicketNum

TicketNum=Request("Num") 
strCallServicedBy = Request.Form("Call_Serviced_By")
strUFName=Request.Form("UFName")	
strULName=Request.Form("ULName")	
	
Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
objConn.Open "DSN=HelpDesk2"
objRec.Open "Calls", objConn

sql = "UPDATE Calls SET  Call_Serviced_By = '" & strCallServicedBy & "' WHERE Ticket_Number=" & Request("Num") & ";"

objConn.Execute(sql)
	
objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing


IF Request.Form("strWorker") <> strCallServicedBy and Session("WorkID") <> strCallServicedBy then

	Set objConn=Server.CreateObject("ADODB.Connection")
	objConn.Open "DSN=Helpdesk2"
	Set objRs=objConn.Execute("Select * from UserInfo WHERE WorkerID = '"& strCallServicedBy &"'")
	Set objRs2=objConn.Execute("Select * from Domains WHERE Country = '"& objRs("Country") &"'")
	Set objRs3=objConn.Execute("Select * from Domains WHERE Country = '"& Session("Country") &"'")
	strSend =objRs("Email") & objRs2("Domain")
	strFromEmail = Session("Email") & objRs3("Domain")
	
	strBody="Help Desk ticket #" + TicketNum + " has been assigned to you by " + Session("FirstName") + " " + Session("LastName") + "." + vbcrlf + vbcrlf +"User: " + strUFName + " " + strULName + vbcrlf + "The problem description for this ticket can be seen online at this address: http://172.16.3.10/modifyticket.asp?Num=" + TicketNum + "."
	objRs.close
	objRs2.close
	objRs3.close
	objConn.Close
	Set objRs = Nothing
	Set objConn = Nothing
	Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
    
	' Set the properties of the object
	objCDOMail.From = strFromEmail '"support@purefishing.com"
	objCDOMail.To = strSend
	objCDOMail.Subject = "You Have a New Ticket"
	objCDOMail.Body = strBody

	objCDOMail.Send

	Set objCDOMail = Nothing  	

End if

Response.Redirect Request.ServerVariables("HTTP_Referer")

%>

