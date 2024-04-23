<%

'Declare local variables to hold the data from the Input form page that is used above.

Dim strFrom
Dim strTo
Dim strDay
Dim strSubmitterName
Dim strFirstName
Dim strMiddleInitial
Dim strLastName
Dim strDepartment
Dim strTitle
Dim strLocation
Dim strSupervisor
Dim strEndDate
Dim strITEndDate
Dim strNotes
Dim objCDOMail 'The CDO object

'First we'll read in the values entered from the form into the Local variables
strFrom = Request.Form("From")
strTo = Request.Form("to")
strDay = Request.Form("day")
strSubmitterName = Request.Form("submittername")
strFirstName = Request.Form("firstname")
strMiddleInitial = Request.Form("middleinitial")
strLastName = Request.Form("lastname")
strDepartment = Request.Form("department")
strTitle = Request.Form("title")
strLocation = Request.Form("location")
strSupervisor = Request.Form("supervisor")
strEndDate = Request.Form("enddate")
strITEndDate = Request.Form("ITenddate")
strNotes = Request.Form("notes")

If strDay = "" Then
	Response.Redirect "errors.asp?Message=1"
End If
If strSubmitterName = "" Then
	Response.Redirect "errors.asp?Message=1"
End If
If strFirstName = "" Then
	Response.Redirect "errors.asp?Message=1"
End If
If strMiddleInitial = "" Then
	Response.Redirect "errors.asp?Message=1"
End If
If strLastName = "" Then
	Response.Redirect "errors.asp?Message=1"
End If
If strDepartment = "" Then
	Response.Redirect "errors.asp?Message=1"
End If
If strEndDate = "" Then
	Response.Redirect "errors.asp?Message=1"
End If

' Create an instance of the NewMail object.
'Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
    
' Set the properties of the object
'objCDOMail.From = StrFrom
'objCDOMail.To = strTo
'objCDOMail.Subject = "Employee Termination Information"
'objCDOMail.Body = "Date: " + strDay + vbcrlf + "Submitted by: " + strSubmitterName + vbcrlf + "(Ex-)Employee Name: " + strFirstName + " " + strMiddleInitial + ". " + strLastName + vbcrlf + "Department: " + strDepartment + vbcrlf + "Title: " + strTitle + vbcrlf + "Location: " + strLocation + vbcrlf + "Supervisor/Manager: " + strSupervisor + vbcrlf + "Last Day of Employment: " + strEndDate + vbcrlf + "Additional Notes: " + strNotes

' There are lots of other properties you can use.
' You can send HTML e-mail, attachments, etc...
' You can also modify most aspects of the message
' like importance, custom headers, ...
' Check the help files for a full list as well
' and the correct syntax.

' Some of the more useful ones I've included samples of here:
' objCDOMail.Cc = "support@purefishing.com"   Notice this sending to more than one person!
'objCDOMail.Bcc = "sschofield@aspfree.com;steve@aspfree.com"
'objCDOMail.Importance = 2 '(0=Low, 1=Normal, 2=High)im a 
'objCDOMail.AttachFile "c:\path\filename.txt", "filename.txt"

' Send the message!
'objCDOMail.Send

' Set the object to nothing because it immediately becomes
' invalid after calling the Send method + it clears it out of the Server's Memory.
'Set objCDOMail = Nothing    

'Response.Redirect ("errors.asp?Message=2")



strDate = FormatDateTime (Date, vbShortDate)
strTime = FormatDateTime (Time, vbShortTime)
strCall_Type = "REQUEST"
strFirstName = Session("FirstName")
strLastName = Replace(Session("LastName"), "'", "''")
strEmpFirstName = Request.Form("FirstName")
strEmpLastName = Replace(Request.Form("LastName"), "'", "''")
strProblem = Replace(Request.form("problem"), "'", "''")
strPackage_Name = Request.Form("Package_Name")
strWorker = Request.Form("Call_Serviced_By")
strExtention = Session("Extention")
strPriority = Request.Form("Priority")
strCountry = Session("Country")
strEmail=Session("Email")
Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
strBody = "Date: " + strDay + vbcrlf + "Submitted by: " + strSubmitterName + vbcrlf + "(Ex-)Employee Name: " + strEmpFirstName + " " + strMiddleInitial + ". " + strEmpLastName + vbcrlf + "Department: " + strDepartment + vbcrlf + "Title: " + strTitle + vbcrlf + "Location: " + strLocation + vbcrlf + "Supervisor/Manager: " + strSupervisor + vbcrlf + "Last Day of Employment: " + strEndDate + vbcrlf +"Date to Terminate IT Services: " + strITEndDate + vbcrlf + "Additional Notes: " + strNotes
	
objConn.Open "DSN=HelpDesk2"
objRec.Open "Calls", objConn

sql = "INSERT INTO Calls(Date_Opened, Time_Opened, Call_Type, User_First_Name, User_Last_Name, Problem_Desc, Package_Name, Closed, In_Process, USER_CONTACT_NUMBER, Email, Call_Serviced_By, Priority,Country) VALUES('"& strDate & "','"& strTime & "','REQUEST','"& strFirstName & "','"& strLastName & "','"& strBody & "','Employment Termination','No','MayBe', '"&strExtention & "', '"&strEmail & "', '', '2', '"&strCountry & "');"
		
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
	Response.Redirect ("../main.asp")

%>
