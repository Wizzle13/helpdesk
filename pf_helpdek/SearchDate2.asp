<HTML>
<BODY>
<%

startdate= request("StartMonth") & "/" & request("StartDay") & "/" &request("StartYear")
Enddate= request("EndMonth") & "/" & request("EndDay") & "/" &request("EndYear")
 
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

If Request("Country") = "All" then
	Set objRs=objConn.Execute("SELECT * FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # )) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # ))  Order by Date_Opened, Date_Closed;")
	Set Count=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # )) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # )) ;")
	Set HCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" #  And (Calls.Call_Type)='Help')) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" #  And (Calls.Call_Type)='Help')) ;")
	Set RCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" #  And (Calls.Call_Type)='Request')) OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" #  And (Calls.Call_Type)='Request')) ;")
	Set OpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # ))  ;")
	Set CloseCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # )) ;")
	Set SOpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # )) And (((Calls.Date_Closed)=#09/22/78#)) ;")
	Set HSOpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # And (Calls.Call_Type)='Help' )) And (((Calls.Date_Closed)=#09/22/78#)) ;")
	Set RSOpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # And (Calls.Call_Type)='Request' )) And (((Calls.Date_Closed)=#09/22/78#)) ;")
Else
	Set objRs=objConn.Execute("SELECT * FROM Calls WHERE ((((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # )) and (Country ='"& Request.form("Country")&"') OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # )) and (Country ='"& Request.form("Country")&"'))  Order by Date_Opened, Date_Closed;")
	Set Count=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # )) and (Country ='"& Request.form("Country")&"') OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # )) and (Country ='"& Request.form("Country")&"')) ;")
	Set HCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" #  And (Calls.Call_Type)='Help')) and (Country ='"& Request.form("Country")&"') OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" #  And (Calls.Call_Type)='Help')) and (Country ='"& Request.form("Country")&"')) ;")
	Set RCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" #  And (Calls.Call_Type)='Request')) and (Country ='"& Request.form("Country")&"') OR (((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" #  And (Calls.Call_Type)='Request')) and (Country ='"& Request.form("Country")&"')) ;")
	Set OpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # )) and (Country ='"& Request.form("Country")&"')) ;")
	Set CloseCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Closed)>=#" & StartDate &" # And (Calls.Date_Closed)<=#" & EndDate &" # )) and (Country ='"& Request.form("Country")&"')) ;")
	Set SOpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # )) And (((Calls.Date_Closed)=#09/22/78#)) and (Country ='"& Request.form("Country")&"')) ;")
	Set HSOpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # And (Calls.Call_Type)='Help' )) And (((Calls.Date_Closed)=#09/22/78#)) and (Country ='"& Request.form("Country")&"')) ;")
	Set RSOpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE ((((Calls.Date_Opened)>=#" & StartDate &" # And (Calls.Date_Opened)<=#" & EndDate &" # And (Calls.Call_Type)='Request' )) And (((Calls.Date_Closed)=#09/22/78#)) and (Country ='"& Request.form("Country")&"')) ;")
End IF

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
	
	Response.Write "<BR style='page-break-after:always'>"
	Response.Write "<table border=0><TR><TD bgcolor=#A8EAA8>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Working on</TD><TD>&nbsp;&nbsp;</TD>"
	Response.Write "<TD bgcolor=#FBFD04>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Pending</TD><TD>&nbsp;&nbsp;</TD>"
	Response.Write "<TD bgcolor=#FF6666>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Stopped</TD></TR></TABLE>"
	Response.Write "<table border=1 cellpadding=0 cellborder=0>"
	Response.Write "<tr><td><B>Ticket #</td>"
	Response.Write "<td><B>User<br>Name</td>"
	Response.Write "<td><B>Call<br>Type</td>"
	Response.Write "<td><B>Date<BR>Opened</td>"
	Response.Write "<td><B>Date<BR>Closed</td>"
	Response.Write "<td><B>Serviced<br>By</td>"
	Response.Write "<td width= 50><B>Problem</td>"

	While Not objRs.EOF
	If objRs("Closed") = "Yes" Then
			Response.Write "<tr bgcolor=lightgrey>"
		Else
			Response.Write "<tr>"
		End If
		Response.Write "<td align=center><a href=""modifyticket.asp?Num=" & objRs("Ticket_Number") & """>" & objRs("Ticket_Number") & "</a>" & "</td>"
		Response.Write "<td>" & objRs("User_First_Name") &"<br>"& objRs("User_Last_Name") & "</td>"	
		If objRs("IN_PROCESS") = "Yes" Then
			Response.Write "<td bgcolor=#A8EAA8 align=center><font color=black>"& objRs("Call_Type") & "</td>"
		End If
		If objRs("IN_PROCESS") = "MayBe" Then
			Response.Write "<td bgcolor=#FBFD04 align=center><font color=black>"& objRs("Call_Type") & "</td>"
		End If

		If objRs("IN_PROCESS") = "No" Then
			Response.Write "<td bgcolor=#FF6666 align=center><font color=black>"& objRs("Call_Type") & "</td>"
		End If
		Response.Write "<td>" & objRs("Date_Opened") & "<br>"& objRs("Time_Opened") &"</td>"
		IF objRs("Date_Closed")="9/22/78" then
			Response.Write "<td><br></td>"
		Else
			Response.Write "<td>" & objRs("Date_Closed") & "<br>"& objRs("Time_Closed") &"</td>"
		End if
		Response.Write "<td>" & objRs("CALL_SERVICED_BY") & "</td>"						
		
		ProblemDesc = objRs("PROBLEM_DESC")
		If Len(ProblemDesc) > 100 Then						
			Response.Write "<td>" & Mid(ProblemDesc, 1, 100) & "...<a href=""modifyticket.asp?Num=" & objRs("Ticket_Number") & """>(more)</a></td>"		
		Else
			Response.Write "<td>" & ProblemDesc & "</td>"
		End If

		
		'************************************
		'*  Fill in the Development Hours column, even if there isn't one in the database
		'************************************
		
		objRs.MoveNext
	Wend
	
	
	
	
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</span>"
objConn.Close
set objConn=nothing


%>

</BODY>
</HTML>