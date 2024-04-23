<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Average Help Ticket Close Time</title>
</head>
<body>
<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
	StartMonth = 1
	StartDay = 1
	StartYear = 2005
	EndMonth = 12
	EndDay = 31
	EndYear = 2005
	Start_Date = StartMonth & "/" & StartDay & "/" & StartYear
	End_Date = EndMonth & "/" & EndDay & "/" & EndYear
	Total = "0"
	Pack_Name = "SAP"
 	Set objRs=objConn.Execute("SELECT * FROM Calls WHERE (CALL_SERVICED_BY = 'RV' or CALL_SERVICED_BY = 'CA' or CALL_SERVICED_BY = 'CG' or CALL_SERVICED_BY = 'CB' or CALL_SERVICED_BY = 'JL' or CALL_SERVICED_BY = 'NS' or CALL_SERVICED_BY = 'ML') and (Date_Opened >= #" & Start_Date & "# and (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and Package_Name Not Like '%" & Pack_Name &"%'")
	Set Count=objConn.Execute("SELECT Count(*) FROM Calls WHERE (CALL_SERVICED_BY = 'RV' or CALL_SERVICED_BY = 'CA' or CALL_SERVICED_BY = 'CG' or CALL_SERVICED_BY = 'CB' or CALL_SERVICED_BY = 'JL' or CALL_SERVICED_BY = 'NS' or CALL_SERVICED_BY = 'ML') and (Date_Opened >= #" & Start_Date &"# And (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and Package_Name Not Like '%" & Pack_Name &"%'")
	
	Response.Write " " & Start_Date & " -" & End_Date & " (" & Count(0) & " Tickets)"
	
	Response.Write "<Table border=1 align=center>"
	Response.Write "<tr>"
	Response.Write "<td Colspan=4><center><h1>Pure Fishing Help Desk Tickets</h1></center></td> "
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td><center>Ticket #</center></td> "
	Response.Write "<td><center>Package Name</center></td> "
	Response.Write "<td><center>Date/Time</center></td> "
	Response.Write "<td><center>Time Open</center></td> "
	Response.Write "<td><center>Running Total</center></td> "
	Response.Write "</tr>"

	While Not objRs.EOF
		TimeOpened = CDate(objRs("Time_opened"))
		TimeClosed = CDate(objRs("Time_Closed"))
		DateOpened = CDate(objRs("Date_Opened"))
		DateClosed = CDate(objRs("Date_Closed"))

		TotalDays = DateDiff("d", DateOpened, DateClosed)

		If TimeOpened > TimeClosed Then
			TotalHours = 24 - (TimeOpened - TimeClosed)
			If DateClosed <> DateOpened Then
				TotalDays = TotalDays - 1
			End If
		Else
			TotalHours = TimeClosed - TimeOpened
		End If

		TotalDaysCumm = TotalDaysCumm + TotalDays
		TotalHoursCumm = TotalHoursCumm + TotalHours
		
		'Adds one day to the day count (Maybe)
		IF TotalHoursCumm > 24 then
			TotalDaysCumm = TotalDaysCumm + 1
			TotalHoursCumm = TotalHoursCumm - 24
		End if

		Response.Write "<tr><td Colspan=1><center>" & objRs("Ticket_Num") & "</center></td> "
		Response.Write "<td Colspan=1><center>" & objRs("Package_Name") & "</center></td> "
		Response.Write "<td Colspan=1><center>" & FormatDateTime(TimeOpened, vbShortTime) & " " & DateOpened & " - " & FormatDateTime(TimeClosed, vbShortTime) & " " & DateClosed & "</center></td>"
		Response.Write "<td Colspan=1><center>" & TotalDays & " Days "& FormatDateTime(TotalHours, vbShortTime) & " Hours </center></td>"
		Response.Write "<td Colspan=1><center>" & TotalDaysCumm & " Days " & FormatDateTime(TotalHoursCumm, vbShortTime) & " Hours </center></td>"
		Response.Write "</tr><tr>"
	
		objRs.MoveNext
	Wend

	Response.Write "</table>"
	Response.Write " "  & TotalDaysCumm &  " Days / " & Count(0) & " Tickets = "& Int(TotalDaysCumm / Count(0))&" Days  " & FormatDateTime( TotalDaysCumm / Count(0), vbShortTime )&" Hours<p>"
	
	Response.Write " "  & FormatDateTime(TotalHoursCumm, vbShortTime) &  " Hours / " & Count(0) & " Tickets = "& FormatDateTime( TotalHoursCumm / Count(0), vbShortTime ) &"<p>"

	Response.Write " "& Int(TotalDaysCumm / Count(0))& " + "  & FormatDateTime( TotalDaysCumm / Count(0), vbShortTime ) &  " + " & FormatDateTime( TotalHoursCumm / Count(0), vbShortTime ) & "  = "& Int(TotalDaysCumm / Count(0))&" Days "&  FormatDateTime((TotalDaysCumm / Count(0)) + ( TotalHoursCumm / Count(0)), vbShortTime)  &" Hours<p>"
%>
</body>
</html>

