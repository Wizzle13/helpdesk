<%
Dim strDate
Dim strTime
Dim strCall_Type
Dim strFirstName
Dim strLastName
Dim strProblem
Dim strPackage_Name
Dim strOpenedBy
Dim strWorker
Dim strExtention
Dim strPriority
Dim strCountry
Dim strEmail
Dim strUpload_File_Location
Dim strTickpriority
Dim objCDOMail
Dim objConn
Dim objRec
Dim sql
Dim objIdentity
Dim strSend
Dim strBody
Dim ObjRs
Dim strTicketNumber
Dim strDomain
Dim ObjRs2
Dim ObjRs3
Dim strFromEmail
Dim strFromCountry
Dim strFrom
Dim Uploader, File
Dim Transport_Date


strDate = Request.Form("Date")
strTime = Request.Form("Time")
strCall_Type = Request.Form("Call_Type")
strFirstName = Request.Form("FirstName")
strLastName = Replace(Request.Form("LastName"), "'", "''")
strProblem = Replace(Request.form("problem"), "'", "''")
strPackage_Name = Request.Form("Package_Name")
strOpenedBy = Session("FirstName") + " " + Session("LastName")
strOpenedBy = Replace(strOpenedBy, "'", "''")
strWorker = Request.Form("Call_Serviced_By")
strExtention = Request.Form("Extention")
strPriority = Request.Form("Priority")
strCountry = Request.Form("Country")
strEmail=Request.Form("Logon_Name")
strScreenShot=Request.Form("ScreenShot")
strUpgrade=Request.Form("Upgrade")
strSAP_Module=Request.Form("SAP_Module")
strTickpriority=Request.Form("tickpriority")
strTransportDate=Request.Form("SAP_Transport_Date")
strUpload_File_Location=Request.Form("Upload_File_Location")
Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")

If Request.Form("Projected_Complete_Date") = "" then
	strProjected_Complete_Date= Request.Form("Projected_Complete_Date2")
Else
	strProjected_Complete_Date= Request.Form("Projected_Complete_Date")
end if
	
objConn.Open "DSN=HelpDesk2"
objRec.Open "Calls", objConn

sql = "INSERT INTO Calls(Date_Opened, Time_Opened, Call_Type, User_First_Name, User_Last_Name, Problem_Desc, Package_Name, Closed, In_Process, USER_CONTACT_NUMBER, Email, Call_Serviced_By, Priority,Country,OpenedBy,Upgrade, SAPPriority, File_Upload_Location, SAP_Module, TransportDate,Projected_Complete_Date) VALUES('"& strDate & "','"& strTime & "','"& strCall_Type & "','"& strFirstName & "','"& strLastName & "','"& strProblem & "','"& strPackage_Name & "','No','MayBe', '"&strExtention & "', '"&strEmail & "', '"&strWorker & "', '"&strPriority & "', '"&strCountry & "','"&strOpenedBy&"','"&strUpgrade&"','"&strTickpriority&"','"&strUpload_File_Location&"','" & strSAP_Module &"','" & strTransportDate &"', '" & strProjected_Complete_Date & "') ;"
		
objConn.Execute(sql)

	
objRec.close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing




	'********Gets Users Last Ticket Number************
	Set ObjConn=Server.CreateObject("ADODB.Connection")
	ObjConn.Open "DSN=HelpDesk2"
	Set ObjRs=ObjConn.Execute("Select * From Calls where eMail = '"& StrEmail &"' Order by Ticket_Number Desc")
	ObjRs.MoveFirst
	strTicketNumber = ObjRs("Ticket_NUmber")
	'****Converts the Ticket Number in to a String******
	strTicketNumber = CStr(strTicketNumber)
	strProblem = objRs("Problem_desc")
	ObjRs.Close
	ObjConn.Close
	Set ObjRs = Nothing
	Set ObjConn = Nothing


'****** Send E-mail to user.*******
If Session("WorkID")<>"" then
	'******Gets Users Email Info*******
	Set ObjConn=Server.CreateObject("ADODB.Connection")
	ObjConn.Open "DSN=Helpdesk2"
	Set ObjRs = ObjConn.Execute("Select * From Domains where Country = '"& strCountry &"'")
	strDomain = ObjRs("Domain")

	Set objRs2=objConn.Execute("Select * from UserInfo WHERE WorkerID = '"& Request.Form("Call_Serviced_By") &"'")
	strFromEmail=objRs2("Email")
	strFromCountry=objRs2("country")

	Set objRs3 = objConn.Execute("Select * From Domains where Country = '"& Session("Country") &"'")
	strFrom = strFromEmail & objRs3("Domain")

	ObjRs.Close
	ObjRs2.Close
	ObjRs3.Close
	ObjConn.Close
	Set ObjRs = Nothing
	Set objRs2 = Nothing
	Set objRs3 = Nothing
	Set ObjConn = Nothing
	
	'******Prepares eMail to be sent to the User*******
	If Session("eMail")<>strEMail or Session("Country")<>strCountry then
		strSend = strEmail & strDomain
		strBody = "This message has been auto-generated by the Online Help Desk.  Please do not reply to this message.  To provide additional information about your open ticket, please login and update your ticket online at the address below." + vbcrlf + vbcrlf + "**************************************************" + vbcrlf + vbcrlf + "Help Desk ticket #" + strTicketNumber + " has been opened for you by " + Session("FirstName") + " " + Session("LastName") + "." + vbcrlf + vbcrlf + "The problem description for this ticket is:" + vbcrlf + strProblem + vbcrlf + vbcrlf + "For more details, you can view this ticket at the following address:" + vbcrlf + "http://172.16.3.10/viewticket.asp?Num=" + strTicketNumber + "."
		'********Sends Email*********
		Set ObjCDOMail = Server.CreateObject("CDONTS.NewMail")
		' Set the properties of the object
		objCDOMail.From = strFrom '"support@purefishing.com"
		objCDOMail.To = strSend
		objCDOMail.Subject = "New Help Desk Ticket"
		objCDOMail.Body = strBody
		objCDOMail.Send
		Set objCDOMail = Nothing
	End if
End if	
'****Send Email to Worker********
If StrWorker<>"" and Session("WorkID")<>Request.Form("Call_Serviced_By") Then
	'******Gets Worker Email Info*******
	Set ObjConn=Server.CreateObject("ADODB.Connection")
	ObjConn.Open "DSN=Helpdesk2"
	Set objRs=objConn.Execute("Select * from UserInfo WHERE WorkerID = '"& Request.Form("Call_Serviced_By") &"'")
	Set objRs2=objConn.Execute("Select * from Domains WHERE Country = '"& objRs("Country") &"'")
	Set objRs3=objConn.Execute("Select * from Domains WHERE Country = '"& Session("Country") &"'")
	strSend =objRs("Email") & objRs2("Domain")
	strFromEmail = Session("Email") & objRs3("Domain")
	ObjRs.Close
	ObjRs2.Close
	objRs3.close
	ObjConn.Close
	Set ObjRs = Nothing
	Set ObjRs2 = Nothing
	Set ObjConn = Nothing
	
	'******Prepares eMail ot be sent to the Worker*******
	strBody="Help Desk ticket #" + strTicketNumber + " has been assigned to you by " + Session("FirstName") + " " + Session("LastName") + "." + vbcrlf + vbcrlf +"User: " + strFirstName + " " + strLastName + vbcrlf + + vbcrlf + vbcrlf + "Problem Description: " + strProblem + vbcrlf + "You can update this ticket at: http://172.16.3.10/modifyticket.asp?Num=" + TicketNum + "."
	'********Sends Email*********
	Set ObjCDOMail = Server.CreateObject("CDONTS.NewMail")
	' Set the properties of the object
	objCDOMail.From = strFromEmail '"support@purefishing.com"
	objCDOMail.To = strSend
	objCDOMail.Subject = "New Help Desk Ticket"
	objCDOMail.Body = strBody
	IF strPriority = 1 then
		objCDOMail.Importance = 2
	End if
	IF strPriority = 2 then
		objCDOMail.Importance = 1
	End if
	IF strPriority = 3 then
		objCDOMail.Importance = 0
	End if
	objCDOMail.Send
	Set objCDOMail = Nothing	
End if	

'****Send Email to HelpDesk********
If Request.Form("Call_Serviced_By")="" Then
	'******Prepares eMail ot be sent to the Worker*******
	strBody="Help Desk ticket #" + strTicketNumber + " has been opened by " + Session("FirstName") + " " + Session("LastName") + ", and has not been assigned." + vbcrlf + vbcrlf + "Problem description for this ticket is:" + vbcrlf + strProblem + vbcrlf + vbcrlf + "You can see this ticket at this address: http://172.16.3.10/modifyticket.asp?Num=" + strTicketNumber + "."
	'********Sends Email*********
	Set ObjCDOMail = Server.CreateObject("CDONTS.NewMail")
	' Set the properties of the object
	objCDOMail.From = "support@purefishing.com"
	objCDOMail.To = "support@purefishing.com"
	objCDOMail.Subject = "New Help Desk Ticket"
	objCDOMail.Body = strBody
	IF strPriority = 1 then
		objCDOMail.Importance = 2
	End if
	IF strPriority = 2 then
		objCDOMail.Importance = 1
	End if
	IF strPriority = 3 then
		objCDOMail.Importance = 0
	End if
	objCDOMail.Send
	Set objCDOMail = Nothing	
End if

Select Case Request.Form("B1")
	Case "Send"
		Response.Redirect "main.asp"
	Case "Save"
		Response.Redirect "main.asp"
	Case "Save & New"
		Response.Redirect "add.asp"
	Case "Save & View"
		Response.Redirect "modifyticket.asp?Num=" + strTicketNumber
End Select

%>

