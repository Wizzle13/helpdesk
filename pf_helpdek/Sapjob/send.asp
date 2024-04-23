<%
dim objCDOMail
dim strBody
'StrComment=Replace(Session("Comments"),"'","''")
StrComment=Session("Comments")
' Create a variable with the strings to display the objCDOMail.Body.
strBody ="Submitted By: " + Session("Email") + vbcrlf +  "Job Name: " + Session("JobName") + vbcrlf + "Job Class: " + Session("JobClass") + vbcrlf + "Target Host: " + Session("TargetHost") + vbcrlf + "Functional Area: " + Session("FunctionalArea") + vbcrlf + "Step Type: " + Session("StepType") + vbcrlf + "Start Time: " + Session("StartTime") + vbcrlf 

Select Case Session("StepType")
	Case "ABAP Program"
		strBody = strBody + "Program Name 1: " + Session("ProgramName1") + vbcrlf + "Variant Name 1: " + Session("VariantName1") + vbcrlf + "Program Name 2: " + Session("ProgramName2") + vbcrlf + "Variant Name 2: " + Session("VariantName2") + vbcrlf + "Program Name 3: " + Session("ProgramName3") + vbcrlf + "Variant Name 3: " + Session("VariantName3") + vbcrlf + "Program Name 4: " + Session("ProgramName4") + vbcrlf + "Variant Name 4: " + Session("VariantName4") + vbcrlf + "Program Name 5: " + Session("ProgramName5") + vbcrlf + "Variant Name 5: " + Session("VariantName5") + vbcrlf + "Language: " + Session("Language") + vbcrlf
	Case "External Command"
		strBody = strBody + "Command Name: " + Session("CommandName") + vbcrlf + "Parameters: " + Session("Parameters") + vbcrlf + "Operating System: " + Session("OperatingSystem") + vbcrlf
	Case "External Program"
		strBody = strBody + "Program Name: " + Session("ProgramName") + vbcrlf + "Parameter: " + Session("Parameter") + vbcrlf

End Select
Select Case Session("StartTime")
	Case "Immediate"
		strBody = strBody +"Start Time:" + Session("StartTime") + "GMT" + vbcrlf		
	Case "Date/Time"
		strBody = strBody + "Start Time:" + Session("StartTime") + vbcrlf + "Date Needed:" +" " + Session("StartHour")+" " + Session("StartMonth")+" " + Session("StartDate") + "GMT" + vbcrlf +"No Start After:" +" "+ Session("NoStartHour")+" " + Session("NoStartMonth")+" " + Session("NoStartDate") + vbcrlf + "Run Time: " + Session("RunTime") + vbcrlf
	Select Case Session("RunTime")
		Case "PeriodicJob"
			strBody = strBody + "Periodic Values: " + Session("PeriodicValues") + vbcrlf
		Select Case Session("PeriodicValues")
			Case "Other Period"
				strBody = strBody + "Other Period: " + Session("OtherPeriod") + vbcrlf + Session("OtherPeriod")+ ": " + Session("TimeSpan") + vbcrlf
		End Select		
	End Select	
	Select Case Session("RunTime")
		Case "PeriodicJob, Restrictions"
			strBody = strBody + "Periodic Values: " + Session("PeriodicValues") + vbcrlf
		Select Case Session("PeriodicValues")
			Case "Other Period"
				strBody = strBody + "Other Period: " + Session("OtherPeriod") + vbcrlf + Session("OtherPeriod")+ ": " + Session("TimeSpan") + vbcrlf
		End Select		
			strBody = strBody + "Restrictions: " + Session("Restrictions") + vbcrlf + "Behavior Restriction: " + Session("BehaviorRestriction") + vbcrlf + "Factory Calander ID: " + Session("Factory Calander ID") + vbcrlf
			
	End Select	
	Case "After Job"	
		strBody = strBody + "Start Time:" + Session("StartTime") + vbcrlf + "After Job: " + Session("AfterJob")	+ vbcrlf	
	Case "After Event"	
		strBody = strBody + "Start Time:" + Session("StartTime") + vbcrlf + "Event Name: " + Session("EventName") + vbcrlf + "Run Time: " + Session("RunTime") + vbcrlf	
		Select Case Session("RunTime")
		Case "PeriodicJob"
			strBody = strBody + "Periodic Values: " + Session("PeriodicValues") + vbcrlf
		Select Case Session("PeriodicValues")
			Case "Other Period"
				strBody = strBody + "Other Period: " + Session("OtherPeriod") + vbcrlf + Session("OtherPeriod")+ ": " + Session("TimeSpan") + vbcrlf
		End Select		
		End Select	
	Case "At Operation Mode"	
		strBody = strBody + "Start Time:" + Session("StartTime") + vbcrlf + "Mode Name:" + Session("ModeName") + vbcrlf
End Select
strBody = strBody + "Recipient Name: " + Session("RecipientName") + vbcrlf + "Type of Output Desired: " + Session("OutputType") + vbcrlf + "Additional Comments: " + Replace(Session("Comments"),"'","''") + vbcrlf

strDate = FormatDateTime (Date, vbShortDate)
strTime = FormatDateTime (Time, vbShortTime)
strCall_Type = "REQUEST"
strFirstName = Session("FirstName")
strLastName = Replace(Session("LastName"), "'", "''")
strProblem = Replace(Request.form("problem"), "'", "''")
strPackage_Name = Request.Form("Package_Name")
strWorker = Request.Form("Call_Serviced_By")
strExtention = Session("Extention")
strPriority = Request.Form("Priority")
strCountry = Session("Country")
strEmail=Session("Email")
Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
	
objConn.Open "DSN=HelpDesk2"
objRec.Open "Calls", objConn

sql = "INSERT INTO Calls(Date_Opened, Time_Opened, Call_Type, User_First_Name, User_Last_Name, Problem_Desc, Package_Name, Closed, In_Process, USER_CONTACT_NUMBER, Email, Call_Serviced_By, Priority,Country) VALUES('"& strDate & "','"& strTime & "','REQUEST','"& strFirstName & "','"& strLastName & "','"& strBody & "','SAP - Job Request','No','MayBe', '"&strExtention & "', '"&strEmail & "', '', '2', '"&strCountry & "');"
		
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