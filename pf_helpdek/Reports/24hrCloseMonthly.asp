<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Percent Closed within 24 hrs</title>
</head>
<body>

<%
If Request("StatYear") = "" then
	StrYear =2005
Else
	StrYear=Request("StatYear")
End if	

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
	StartMonth = 1
	StartDay = 1
	StartYear = StrYear
	EndMonth = 1
	EndDay = 31
	EndYear = StrYear
	Start_Date = StartMonth & "/" & StartDay & "/" & StartYear
	End_Date = EndMonth & "/" & EndDay & "/" & EndYear
	Total = "0"
	Pack_Name = "SAP"
 	Set objRs=objConn.Execute("SELECT * FROM Calls WHERE (CALL_SERVICED_BY = 'RV' or CALL_SERVICED_BY = 'CH' or CALL_SERVICED_BY = 'CB' or CALL_SERVICED_BY = 'JL' or CALL_SERVICED_BY = 'NS') and (Date_Opened >= #" & Start_Date & "# and (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and (Package_Name Not Like '%" & Pack_Name &"%' or Package_Name = 'SAP - Unlock User')")
	Set Count=objConn.Execute("SELECT Count(*) FROM Calls WHERE (CALL_SERVICED_BY = 'RV' or CALL_SERVICED_BY = 'CH' or CALL_SERVICED_BY = 'CB' or CALL_SERVICED_BY = 'JL' or CALL_SERVICED_BY = 'NS') and (Date_Opened >= #" & Start_Date &"# And (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and (Package_Name Not Like '%" & Pack_Name &"%' or Package_Name = 'SAP - Unlock User')")
	
	Response.Write "<center>" & Start_Date & " -" & End_Date & " (" & Count(0) & " Tickets)</center>"
	
	
	While Not objRs.EOF
		TimeOpened = CDate(objRs("Time_opened"))
		TimeClosed = CDate(objRs("Time_Closed"))
		DateOpened = CDate(objRs("Date_Opened"))
		DateClosed = CDate(objRs("Date_Closed"))

		Hours_Open=(((DATECLOSED-DATEOPENED)*1440)+(TIMECLOSED-TIMEOPENED)*1440)/60
		
		'Adds one day to the day count (Maybe)
		IF Hours_Open >= 24 then
			Over = Over + 1
		Else
			Under = Under + 1
		End if		
		objRs.MoveNext
	Wend

	
	
	Response.Write "<p><center> "
	Response.Write "<p> "& Int(Over)&" Over "& Int(Under)&" Under<p>"
	Response.Write "<p> "& Round((Under / Count(0).Value)*100, 2)&"% <p>"

	%>
</form>
</body>
</html>

