<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Percent Closed within 24 hrs</title>
</head>
<body>

<%
If Request("StatYear") = "" then
	StrYear =2006
Else
	StrYear=Request("StatYear")
End if	

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
	StartMonth = 1
	StartDay = 1
	StartYear = StrYear
	EndMonth = 12
	EndDay = 31
	EndYear = StrYear
	Start_Date = StartMonth & "/" & StartDay & "/" & StartYear
	End_Date = EndMonth & "/" & EndDay & "/" & EndYear
	TotalRV = "0"
	Pack_Name = "SAP"
 	Set objRsRV=objConn.Execute("SELECT * FROM Calls WHERE (CALL_SERVICED_BY = 'RV') and (Date_Opened >= #" & Start_Date & "# and (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and (Package_Name Not Like '%" & Pack_Name &"%' or Package_Name = 'SAP - Unlock User')")
	Set CountRV=objConn.Execute("SELECT Count(*) FROM Calls WHERE (CALL_SERVICED_BY = 'RV') and (Date_Opened >= #" & Start_Date &"# And (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and (Package_Name Not Like '%" & Pack_Name &"%' or Package_Name = 'SAP - Unlock User')")
	
		
	While Not objRsRV.EOF
		TimeOpened = CDate(objRsRV("Time_opened"))
		TimeClosed = CDate(objRsRV("Time_Closed"))
		DateOpened = CDate(objRsRV("Date_Opened"))
		DateClosed = CDate(objRsRV("Date_Closed"))

		Hours_Open=(((DATECLOSED-DATEOPENED)*1440)+(TIMECLOSED-TIMEOPENED)*1440)/60
		
		'Adds one day to the day count (Maybe)
		IF Hours_Open >= 24 then
			OverRV = OverRV + 1
		Else
			UnderRV = UnderRV + 1
		End if
		
		objRsRV.MoveNext
	Wend

	
	
	
	TotalCB = "0"
	Pack_Name = "SAP"
 	Set objRsCB=objConn.Execute("SELECT * FROM Calls WHERE (CALL_SERVICED_BY = 'CB') and (Date_Opened >= #" & Start_Date & "# and (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and (Package_Name Not Like '%" & Pack_Name &"%' or Package_Name = 'SAP - Unlock User')")
	Set CountCB=objConn.Execute("SELECT Count(*) FROM Calls WHERE (CALL_SERVICED_BY = 'CB') and (Date_Opened >= #" & Start_Date &"# And (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and (Package_Name Not Like '%" & Pack_Name &"%' or Package_Name = 'SAP - Unlock User')")
	
	
	
	While Not objRsCB.EOF
		TimeOpened = CDate(objRsCB("Time_opened"))
		TimeClosed = CDate(objRsCB("Time_Closed"))
		DateOpened = CDate(objRsCB("Date_Opened"))
		DateClosed = CDate(objRsCB("Date_Closed"))

		Hours_Open=(((DATECLOSED-DATEOPENED)*1440)+(TIMECLOSED-TIMEOPENED)*1440)/60
		
		'Adds one day to the day count (Maybe)
		IF Hours_Open >= 24 then
			OverCB = OverCB + 1
		Else
			UnderCB = UnderCB + 1
		End if
		objRsCB.MoveNext
	Wend

	
	
	Total2 = "0"
	Pack_Name = "SAP"
 	Set objRsCH=objConn.Execute("SELECT * FROM Calls WHERE (CALL_SERVICED_BY = 'CH') and (Date_Opened >= #" & Start_Date & "# and (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and (Package_Name Not Like '%" & Pack_Name &"%' or Package_Name = 'SAP - Unlock User')")
	Set CountCH=objConn.Execute("SELECT Count(*) FROM Calls WHERE (CALL_SERVICED_BY = 'CH') and (Date_Opened >= #" & Start_Date &"# And (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and (Package_Name Not Like '%" & Pack_Name &"%' or Package_Name = 'SAP - Unlock User')")
	
	
	While Not objRsCH.EOF
		TimeOpened = CDate(objRsCH("Time_opened"))
		TimeClosed = CDate(objRsCH("Time_Closed"))
		DateOpened = CDate(objRsCH("Date_Opened"))
		DateClosed = CDate(objRsCH("Date_Closed"))

		Hours_Open=(((DATECLOSED-DATEOPENED)*1440)+(TIMECLOSED-TIMEOPENED)*1440)/60
		
		'Adds one day to the day count (Maybe)
		IF Hours_Open >= 24 then
			OverCH = OverCH + 1
		Else
			UnderCH = UnderCH + 1
		End if
		objRsCH.MoveNext
	Wend

	
	TotalJL = "0"
	Pack_Name = "SAP"
 	Set objRsJL=objConn.Execute("SELECT * FROM Calls WHERE (CALL_SERVICED_BY = 'JL') and (Date_Opened >= #" & Start_Date & "# and (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and (Package_Name Not Like '%" & Pack_Name &"%' or Package_Name = 'SAP - Unlock User')")
	Set CountJL=objConn.Execute("SELECT Count(*) FROM Calls WHERE (CALL_SERVICED_BY = 'JL') and (Date_Opened >= #" & Start_Date &"# And (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and (Package_Name Not Like '%" & Pack_Name &"%' or Package_Name = 'SAP - Unlock User')")
	
	
	
	While Not objRsJL.EOF
		TimeOpened = CDate(objRsJL("Time_opened"))
		TimeClosed = CDate(objRsJL("Time_Closed"))
		DateOpened = CDate(objRsJL("Date_Opened"))
		DateClosed = CDate(objRsJL("Date_Closed"))

		Hours_Open=(((DATECLOSED-DATEOPENED)*1440)+(TIMECLOSED-TIMEOPENED)*1440)/60
		
		'Adds one day to the day count (Maybe)
		IF Hours_Open >= 24 then
			OverJL = OverJL + 1
		Else
			UnderJL = UnderJL + 1
		End if
		objRsJL.MoveNext
	Wend

	

	
TotalNS = "0"
	Pack_Name = "SAP"
 	Set objRsNS=objConn.Execute("SELECT * FROM Calls WHERE (CALL_SERVICED_BY = 'NS') and (Date_Opened >= #" & Start_Date & "# and (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and (Package_Name Not Like '%" & Pack_Name &"%' or Package_Name = 'SAP - Unlock User')")
	Set CountNS=objConn.Execute("SELECT Count(*) FROM Calls WHERE (CALL_SERVICED_BY = 'NS') and (Date_Opened >= #" & Start_Date &"# And (Date_Closed <= #" & End_Date & "# And Date_Closed >= #" & Start_Date & "#)) and Call_Type='Help' and (Package_Name Not Like '%" & Pack_Name &"%' or Package_Name = 'SAP - Unlock User')")
	
	
	While Not objRsNS.EOF
		TimeOpened = CDate(objRsNS("Time_opened"))
		TimeClosed = CDate(objRsNS("Time_Closed"))
		DateOpened = CDate(objRsNS("Date_Opened"))
		DateClosed = CDate(objRsNS("Date_Closed"))

		Hours_Open=(((DATECLOSED-DATEOPENED)*1440)+(TIMECLOSED-TIMEOPENED)*1440)/60
		
		'Adds one day to the day count (Maybe)
		IF Hours_Open >= 24 then
			OverNS = OverNS + 1
		Else
			UnderNS = UnderNS + 1
		End if
		objRsNS.MoveNext
	Wend

	
	Response.Write "<Table border=1>"
	Response.Write "<tr><td>Ryan</td><td>"& CountRV(0) & " Tickets </td><td>"& Int(OverRV)&" Over </td><td>"& Int(UnderRV)&" Under </td><td>"& Round((UnderRV / CountRV(0).Value)*100, 2)&"% </td></tr>"
	Response.Write "<tr><td>Chris Burton</td><td>"& CountCB(0) & " Tickets </td><td>"& Int(OverCB)&" Over </td><td>"& Int(UnderCB)&" Under </td><td>"& Round((UnderCB / CountCB(0).Value)*100, 2)&"% </td></tr>"
	Response.Write "<tr><td>Chris Hoerich</td><td>"& CountCH(0) & " Tickets </td><td>"& Int(OverCH)&" Over </td><td>"& Int(UnderCH)&" Under </td><td>"& Round((UnderCH / CountCH(0).Value)*100, 2)&"% </td></tr>"
	Response.Write "<tr><td>Josh</td><td>"& CountJL(0) & " Tickets </td><td>"& Int(OverJL)&" Over </td><td>"& Int(UnderJL)&" Under </td><td>"& Round((UnderJL / CountJL(0).Value)*100, 2)&"% </td></tr>"
	Response.Write "<tr><td>Nick</td><td>"& CountNS(0) & " Tickets </td><td>"& Int(OverNS)&" Over </td><td>"& Int(UnderNS)&" Under </td><td>"& Round((UnderNS / CountNS(0).Value)*100, 2)&"% </td></tr>"
	Response.Write "</table>"
	TotalOver= OverRV+OverCB+OverCH+OverJL+OverNS
	TotalUnder= UnderRV+UnderCB+UnderCH+UnderJL+UnderNS
	TotalTickets= CountRV(0)+CountCB(0)+CountCH(0)+CountJL(0)+CountNS(0)
	
	Response.Write "Toatl Over "& Int(TotalOver)
	Response.Write "<p>Toatl Under "& Int(TotalUnder)
	Response.Write "<p>Toatl Tickets "& Int(TotalTickets)	
	Response.Write "<p><p><center><a href='24hrClosebyPerson.asp'>Refresh</a></center>"
	%>
	<center><form method="post" action="24hrClosebyperson.asp" name="24hrClose">
	<%
	Response.Write "<p><p><Select name=StatYear><Option value=1999>1999 <Option value=2000>2000 <Option value=2001>2001 <Option value=2002>2002 <Option value=2003>2003 <Option value=2004>2004<Option value=2005>2005 <Option selected value=2006>2006</select>"
%>
<input type="submit" value="Go" name="B1"></center>
</form>
</body>
</html>

