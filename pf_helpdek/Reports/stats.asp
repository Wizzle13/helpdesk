<HTML>
<head>
<title>Stats</title>
</head>
<BODY>
<%

startdate= request("StartMonth") & "/" & request("StartDay") & "/" &request("StartYear")
Enddate= request("EndMonth") & "/" & request("EndDay") & "/" &request("EndYear")
 
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
'All Stats
	Set objRs=objConn.Execute("SELECT * FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # )) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # ))  Order by Date_Opened, Date_Closed;")
	Set Count=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # )) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # )) ;")
	Set HCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" #  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" #  And (Calls.Call_Type)='Help')) ;")
	Set RCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" #  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" #  And (Calls.Call_Type)='Request')) ;")
	Set TotalOpened=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # )) ;")	
	Set StillOpened=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # And ((Calls.Date_Closed)>#" & EndDate &" # Or (Calls.Closed)='No') ) ;")	

    Total = "0"
	Pack_Name = "SAP"
 	Set objRs=objConn.Execute("SELECT * FROM Calls WHERE (Date_Opened >= #" & StartDate & "# And (Date_Closed <= #" & EndDate & "# And Date_Closed >= #" & StartDate & "#)) and Call_Type='Help' and Package_Name Not Like '%" & Pack_Name &"%'")
	Set AveCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (Date_Opened >= #" & StartDate &"# And (Date_Closed <= #" & EndDate & "# And Date_Closed >= #" & StartDate & "#)) and Call_Type='Help' and Package_Name Not Like '%" & Pack_Name &"%'")
	
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

		
		objRs.MoveNext
	Wend
'North America Stats
	Set objRs1=objConn.Execute("SELECT * FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # and (Country ='US' or Country ='CA'))) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # and (Country ='US' or Country ='CA')))  Order by Date_Opened, Date_Closed;")
	Set objDomains=objConn.Execute("SELECT * FROM Domains WHERE Helpdesk = 'America' ;")
	'Set Count1=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # and (Country ='"& objDomains("Country") &"'))) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # and (Country ='"& objDomains("Country") &"'))) ;")
	Set Count1=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # and (Country ='US' or Country ='CA'))) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # and (Country ='US' or Country ='CA'))) ;")
	Set HCount1=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # and (Country ='US' or Country ='CA') And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # and (Country ='US' or Country ='CA') And (Calls.Call_Type)='Help')) ;")
	Set RCount1=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # and (Country ='US' or Country ='CA') And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # and (Country ='US' or Country ='CA') And (Calls.Call_Type)='Request')) ;")
	Set TotalOpened1=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # ) and (Country ='US' or Country ='CA')) ;")	
	Set StillOpened1=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" #  And ((Calls.Closed)='No' Or (Calls.Date_Closed)>#" & EndDate &" # ) and (Country ='US' or Country ='CA')) ;")	
	 Total1 = "0"
	Pack_Name = "SAP"
 	Set objRs1=objConn.Execute("SELECT * FROM Calls WHERE (Date_Opened >= #" & StartDate & "# And (Date_Closed <= #" & EndDate & "# And Date_Closed >= #" & StartDate & "#)) and (Country ='US' or Country ='CA') and Call_Type='Help' and Package_Name Not Like '%" & Pack_Name &"%'")
	Set AveCount1=objConn.Execute("SELECT Count(*) FROM Calls WHERE (Date_Opened >= #" & StartDate &"# And (Date_Closed <= #" & EndDate & "# And Date_Closed >= #" & StartDate & "#)) and (Country ='US' or Country ='CA') and Call_Type='Help' and Package_Name Not Like '%" & Pack_Name &"%'")
	
	While Not objRs1.EOF
		TimeOpened1 = CDate(objRs1("Time_opened"))
		TimeClosed1 = CDate(objRs1("Time_Closed"))
		DateOpened1 = CDate(objRs1("Date_Opened"))
		DateClosed1 = CDate(objRs1("Date_Closed"))

		TotalDays1 = DateDiff("d", DateOpened1, DateClosed1)

		If TimeOpened1 > TimeClosed1 Then
			TotalHours1 = 24 - (TimeOpened1 - TimeClosed1)
			If DateClosed1 <> DateOpened1 Then
				TotalDays1 = TotalDays1 - 1
			End If
		Else
			TotalHours1 = TimeClosed1 - TimeOpened1
		End If

		TotalDaysCumm1 = TotalDaysCumm1 + TotalDays1
		TotalHoursCumm1 = TotalHoursCumm1 + TotalHours1
		
		'Adds one day to the day count (Maybe)
		IF TotalHoursCumm1 > 24 then
			TotalDaysCumm1 = TotalDaysCumm1 + 1
			TotalHoursCumm1 = TotalHoursCumm1 - 24
		End if

		
		objRs1.MoveNext
	Wend
'Atlantic Rim Stats	
	Set objRs2=objConn.Execute("SELECT * FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # and (Country ='SE' or Country ='DK' or Country ='UK' or Country ='FI' or Country ='NO' or Country ='FR'))) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # and (Country ='SE' or Country ='DK' or Country ='UK' or Country ='FI' or Country ='NO' or Country ='FR')))  Order by Date_Opened, Date_Closed;")
	Set Count2=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # and (Country ='SE' or Country ='DK' or Country ='UK' or Country ='FI' or Country ='NO' or Country ='FR'))) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # and (Country ='SE' or Country ='DK' or Country ='UK' or Country ='FI' or Country ='NO' or Country ='FR'))) ;")
	Set HCount2=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # and (Country ='SE' or Country ='DK' or Country ='UK' or Country ='FI' or Country ='NO' or Country ='FR') And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # and (Country ='SE' or Country ='DK' or Country ='UK' or Country ='FI' or Country ='NO' or Country ='FR') And (Calls.Call_Type)='Help')) ;")
	Set RCount2=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # and (Country ='SE' or Country ='DK' or Country ='UK' or Country ='FI' or Country ='NO' or Country ='FR') And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # and (Country ='SE' or Country ='DK' or Country ='UK' or Country ='FI' or Country ='NO' or Country ='FR') And (Calls.Call_Type)='Request')) ;")
	Set TotalOpened2=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # ) and (Country ='SE' or Country ='DK' or Country ='UK' or Country ='FI' or Country ='NO' or Country ='FR')) ;")	
	Set StillOpened2=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((Calls.Date_Opened)>=#" & StartDate &"# And (Calls.Date_Opened)<=#" & EndDate &" # And ((Calls.Closed)='No' Or (Calls.Date_Closed)>#" & EndDate &" # ) and (Country ='SE' or Country ='DK' or Country ='UK' or Country ='FI' or Country ='NO' or Country ='FR')) ;")	
	 Total2 = "0"
	Pack_Name = "SAP"
 	Set objRs2=objConn.Execute("SELECT * FROM Calls WHERE (Date_Opened >= #" & StartDate & "# And (Date_Closed <= #" & EndDate & "# And Date_Closed >= #" & StartDate & "#)) and (Country ='SE' or Country ='DK' or Country ='UK' or Country ='FI' or Country ='NO' or Country ='FR') and Call_Type='Help' and Package_Name Not Like '%" & Pack_Name &"%'")
	Set AveCount2=objConn.Execute("SELECT Count(*) FROM Calls WHERE (Date_Opened >= #" & StartDate &"# And (Date_Closed <= #" & EndDate & "# And Date_Closed >= #" & StartDate & "#)) and (Country ='SE' or Country ='DK' or Country ='UK' or Country ='FI' or Country ='NO' or Country ='FR') and Call_Type='Help' and Package_Name Not Like '%" & Pack_Name &"%'")
	
	While Not objRs2.EOF
		TimeOpened2 = CDate(objRs2("Time_opened"))
		TimeClosed2 = CDate(objRs2("Time_Closed"))
		DateOpened2 = CDate(objRs2("Date_Opened"))
		DateClosed2 = CDate(objRs2("Date_Closed"))

		TotalDays2 = DateDiff("d", DateOpened2, DateClosed2)

		If TimeOpened2 > TimeClosed2 Then
			TotalHours2 = 24 - (TimeOpened2 - TimeClosed2)
			If DateClosed2 <> DateOpened2 Then
				TotalDays2 = TotalDays2 - 1
			End If
		Else
			TotalHours2 = TimeClosed2 - TimeOpened2
		End If

		TotalDaysCumm2 = TotalDaysCumm2 + TotalDays2
		TotalHoursCumm2 = TotalHoursCumm2 + TotalHours2
		
		'Adds one day to the day count (Maybe)
		IF TotalHoursCumm2 > 24 then
			TotalDaysCumm2 = TotalDaysCumm2 + 1
			TotalHoursCumm2 = TotalHoursCumm2 - 24
		End if

		
		objRs2.MoveNext
	Wend
'Pacific Rim Stats
	Set objRs3=objConn.Execute("SELECT * FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # and (Country ='CN' or Country ='TW' or Country ='JP' or Country ='AU' or Country ='NZ' or Country ='KR' or Country ='MY' or Country ='TH'))) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # and (Country ='CN' or Country ='TW' or Country ='JP' or Country ='AU' or Country ='NZ' or Country ='KR' or Country ='MY' or Country ='TH')))  Order by Date_Opened, Date_Closed;")
	Set Count3=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # and (Country ='CN' or Country ='TW' or Country ='JP' or Country ='AU' or Country ='NZ' or Country ='KR' or Country ='MY' or Country ='TH'))) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # and (Country ='CN' or Country ='TW' or Country ='JP' or Country ='AU' or Country ='NZ' or Country ='KR' or Country ='MY' or Country ='TH'))) ;")
	Set HCount3=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # and (Country ='CN' or Country ='TW' or Country ='JP' or Country ='AU' or Country ='NZ' or Country ='KR' or Country ='MY' or Country ='TH') And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # and (Country ='CN' or Country ='TW' or Country ='JP' or Country ='AU' or Country ='NZ' or Country ='KR' or Country ='MY' or Country ='TH') And (Calls.Call_Type)='Help')) ;")
	Set RCount3=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # and (Country ='CN' or Country ='TW' or Country ='JP' or Country ='AU' or Country ='NZ' or Country ='KR' or Country ='MY' or Country ='TH') And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # and (Country ='CN' or Country ='TW' or Country ='JP' or Country ='AU' or Country ='NZ' or Country ='KR' or Country ='MY' or Country ='TH') And (Calls.Call_Type)='Request')) ;")
	Set TotalOpened3=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # ) and (Country ='CN' or Country ='TW' or Country ='JP' or Country ='AU' or Country ='NZ' or Country ='KR' or Country ='MY' or Country ='TH')) ;")	
	Set StillOpened3=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" #  And ((Calls.Closed)='No' Or (Calls.Date_Closed)>#" & EndDate &" # ) and (Country ='CN' or Country ='TW' or Country ='JP' or Country ='AU' or Country ='NZ' or Country ='KR' or Country ='MY' or Country ='TH')) ;")	
	 Total3 = "0"
	Pack_Name = "SAP"
 	Set objRs3=objConn.Execute("SELECT * FROM Calls WHERE (Date_Opened >= #" & StartDate & "# And (Date_Closed <= #" & EndDate & "# And Date_Closed >= #" & StartDate & "#)) and (Country ='CN' or Country ='TW' or Country ='JP' or Country ='AU' or Country ='NZ' or Country ='KR' or Country ='MY' or Country ='TH') and Call_Type='Help' and Package_Name Not Like '%" & Pack_Name &"%'")
	Set AveCount3=objConn.Execute("SELECT Count(*) FROM Calls WHERE (Date_Opened >= #" & StartDate &"# And (Date_Closed <= #" & EndDate & "# And Date_Closed >= #" & StartDate & "#)) and (Country ='CN' or Country ='TW' or Country ='JP' or Country ='AU' or Country ='NZ' or Country ='KR' or Country ='MY' or Country ='TH') and Call_Type='Help' and Package_Name Not Like '%" & Pack_Name &"%'")
	
	While Not objRs3.EOF
		TimeOpened3 = CDate(objRs3("Time_opened"))
		TimeClosed3 = CDate(objRs3("Time_Closed"))
		DateOpened3 = CDate(objRs3("Date_Opened"))
		DateClosed3 = CDate(objRs3("Date_Closed"))

		TotalDays3 = DateDiff("d", DateOpened3, DateClosed3)

		If TimeOpened3 > TimeClosed3 Then
			TotalHours3 = 24 - (TimeOpened3 - TimeClosed3)
			If DateClosed3 <> DateOpened3 Then
				TotalDays3 = TotalDays3 - 1
			End If
		Else
			TotalHours3 = TimeClosed3 - TimeOpened3
		End If

		TotalDaysCumm3 = TotalDaysCumm3 + TotalDays3
		TotalHoursCumm3 = TotalHoursCumm3 + TotalHours3
		
		'Adds one day to the day count (Maybe)
		IF TotalHoursCumm3 > 24 then
			TotalDaysCumm3 = TotalDaysCumm3 + 1
			TotalHoursCumm3 = TotalHoursCumm3 - 24
		End if

		
		objRs3.MoveNext
	Wend
	Response.Write "<Table border=1 align=center>"
	Response.Write "<tr>"
	Response.Write "<td Colspan=3><center><h1>Pure Fishing Help Desk Tickets</h1></center></td> "
	Response.Write "</tr><tr>"
	Response.Write "<td Colspan=3><center>" & StartDate & " - " & EndDate &"</center></td> "
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& Count(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Help Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& HCount(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Request Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& RCount(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Tickes Opened in previous Month</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& TotalOpened(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Tickes Still Opened From previous Month</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& StillOpened(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Average Close Time of Help Tickets from previous Month</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& Int(TotalDaysCumm / AveCount(0))&" Days "&  FormatDateTime((TotalDaysCumm / AveCount(0)) + ( TotalHoursCumm / AveCount(0)), vbShortTime)  &" Hours</td>"
	Response.Write "</tr><tr>"
		

	Response.Write "<td Colspan=3><hr></td>"
	Response.Write "</tr><tr>"
	Response.Write "<td Colspan=3>*This Number is the total of Opened and/or Closed Tickets dring the time peroid.</td> "
	Response.Write "</tr>"
	Response.Write "</table>"
	
	Response.Write "<Table border=1 align=center>"
	Response.Write "<tr>"
	Response.Write "<td Colspan=3><center><h1>North America Help Desk Tickets</h1></center></td> "
	Response.Write "</tr><tr>"
	Response.Write "<td Colspan=3><center>" & StartDate & " - " & EndDate &"</center></td> "
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& Count1(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Help Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& HCount1(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Request Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& RCount1(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Tickes Opened in previous</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& TotalOpened1(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Tickes Still Opened from previous</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& StillOpened1(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Average Close Time of Help Tickets from previous</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& Int(TotalDaysCumm1 / AveCount1(0))&" Days "&  FormatDateTime((TotalDaysCumm1 / AveCount1(0)) + ( TotalHoursCumm1 / AveCount1(0)), vbShortTime)  &" Hours</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td Colspan=3><hr></td>"
	Response.Write "</tr><tr>"
	Response.Write "<td Colspan=3>*This Number is the total of Opened and/or Closed Tickets dring the time peroid.</td> "
	Response.Write "</tr>"
	Response.Write "</table>"
	
	Response.Write "<Table border=1 align=center>"
	Response.Write "<tr>"
	Response.Write "<td Colspan=3><center><h1>Atlantic Rim Help Desk Tickets</h1></center></td> "
	Response.Write "</tr><tr>"
	Response.Write "<td Colspan=3><center>" & StartDate & " - " & EndDate &"</center></td> "
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& Count2(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Help Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& HCount2(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Request Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& RCount2(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Tickes Opened from previous Month</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& TotalOpened2(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Tickes Still Opened from previous Month</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& StillOpened2(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Average Close Time of Help Tickets from previous Month</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& Int(TotalDaysCumm2 / AveCount2(0))&" Days "&  FormatDateTime((TotalDaysCumm2 / AveCount2(0)) + ( TotalHoursCumm2 / AveCount2(0)), vbShortTime)  &" Hours</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td Colspan=3><hr></td>"
	Response.Write "</tr><tr>"
	Response.Write "<td Colspan=3>*This Number is the total of Opened and/or Closed Tickets dring the time peroid.</td> "
	Response.Write "</tr>"
	Response.Write "</table>"

	Response.Write "<Table border=1 align=center>"
	Response.Write "<tr>"
	Response.Write "<td Colspan=3><center><h1>Pacific Rim Help Desk Tickets</h1></center></td> "
	Response.Write "</tr><tr>"
	Response.Write "<td Colspan=3><center>" & StartDate & " - " & EndDate &"</center></td> "
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& Count3(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Help Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& HCount3(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Request Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& RCount3(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Tickes Opened from previous Month</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& TotalOpened3(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Tickes Still Opened from previous Month</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& StillOpened3(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Average Close Time of Help Tickets from previous Month</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& Int(TotalDaysCumm3 / AveCount3(0))&" Days "&  FormatDateTime((TotalDaysCumm3 / AveCount3(0)) + ( TotalHoursCumm3 / AveCount3(0)), vbShortTime)  &" Hours</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td Colspan=3><hr></td>"
	Response.Write "</tr><tr>"
	Response.Write "<td Colspan=3>*This Number is the total of Opened and/or Closed Tickets dring the time peroid.</td> "
	Response.Write "</tr>"	
	Response.Write "</table>"


%>

</BODY>
</HTML>