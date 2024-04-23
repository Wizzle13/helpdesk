<HTML>
<BODY>
<%
strMonth= request("StartMonth")
strYear= request("StartYear")
startdate= request("StartMonth") & "/" & request("StartDay") & "/" &request("StartYear")
Enddate= request("EndMonth") & "/" & request("EndDay") & "/" &request("EndYear")
 
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

Set objRs=objConn.Execute("SELECT * FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # )) OR (((Calls.Date_Closed)>#" & StartDate &" # And (Calls.Date_Closed)<#" & EndDate &" # ))  Order by Date_Opened, Date_Closed;")
Set Count=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # )) OR (((Calls.Date_Closed)>#" & StartDate &" # And (Calls.Date_Closed)<#" & EndDate &" # )) ;")
Set HCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" #  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>#" & StartDate &" # And (Calls.Date_Closed)<#" & EndDate &" #  And (Calls.Call_Type)='Help')) ;")
Set RCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" #  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>#" & StartDate &" # And (Calls.Date_Closed)<#" & EndDate &" #  And (Calls.Call_Type)='Request')) ;")
Set OpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # ))  ;")
Set CloseCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Closed)>#" & StartDate &" # And (Calls.Date_Closed)<#" & EndDate &" # )) ;")
Set SOpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # )) And (((Calls.Date_Closed)=#09/22/78#)) ;")
Set HSOpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # And (Calls.Call_Type)='Help' )) And (((Calls.Date_Closed)=#09/22/78#)) ;")
Set RSOpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # And (Calls.Call_Type)='Request' )) And (((Calls.Date_Closed)=#09/22/78#)) ;")
	Response.Write "<Table border=1 align=center>"
	Response.Write "<tr>"
	Response.Write "<td Colspan=3><center><h1>Help Desk Tickets</h1></center></td> "
	Response.Write "</tr><tr>"
	Response.Write "<td Colspan=3><center>" & StartDate & " - " & EndDate &"</center></td> "
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& Count(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Help Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& HCount(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Total # of Request Tickes*</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& RCount(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td Colspan=3><hr></td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Request Tickes Remaining Open</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& RSOpenCount(0) &"</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Help Tickes Remaining Open</td><td>&nbsp;&nbsp;&nbsp;</td><td>"& HSOpenCount(0) &"</td>"	
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td Colspan=3>*This Number is the total of Opened and/or Closed Tickets dring the time peroid.</td> "
	Response.Write "</tr>"
	Response.Write "</table>"

Select Case strMonth
	Case 1
		StrMonth="January"
	Case 2
		StrMonth="Febuary"
	Case 3
		StrMonth="March"
	Case 4
		StrMonth="April"
	Case 5
		StrMonth="May"
	Case 6
		StrMonth="June"
	Case 7
		StrMonth="July"
	Case 8
		StrMonth="August"
	Case 9
		StrMonth="September"
	Case 10
		StrMonth="October"
	Case 11
		StrMonth="November"
	Case 12
		StrMonth="December"
End Select	

	Response.Write "<form method='post' action='send2.asp' name='ViewWork'>"
	Response.Write "<input type=Hidden Name=strMonth Value=" & strMonth &">"
	Response.Write "<input type=Hidden Name=strYear Value=" & strYear &">"
	Response.Write "<input type=Hidden Name=strCount Value=" & Count(0) &">"
	Response.Write "<input type=Hidden Name=strHCount Value=" & HCount(0) &">"	
	Response.Write "<input type=Hidden Name=strRCount Value=" & RCount(0) &">"		
	Response.Write "<input type=Hidden Name=strRSOpenCount Value=" & RSOpenCount(0) &">"		
	Response.Write "<input type=Hidden Name=strHSOpenCount Value=" & HSOpenCount(0) &">"			
	Response.Write "<input type=Submit  value=Save>"
	Response.Write "</form>"
	
	objConn.Close
set objConn=nothing
%>

</BODY>
</HTML>