<!--#include file="_head.asp"-->
<%

startdate= request("StartMonth") & "/" & request("StartDay") & "/" &request("StartYear")
Enddate= request("EndMonth") & "/" & request("EndDay") & "/" &request("EndYear")
 
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

Set objRs=objConn.Execute("SELECT * FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # )) OR (((Calls.Date_Closed)>#" & StartDate &" # And (Calls.Date_Closed)<#" & EndDate &" # ))  Order by Date_Opened, Date_Closed;")
Set Count=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # )) OR (((Calls.Date_Closed)>#" & StartDate &" # And (Calls.Date_Closed)<#" & EndDate &" # )) ;")
Set OpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # ))  ;")
Set CloseCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Closed)>#" & StartDate &" # And (Calls.Date_Closed)<#" & EndDate &" # )) ;")
Set SOpenCount=objConn.Execute("SELECT Count(*) FROM Calls WHERE (((Calls.Date_Opened)>#" & StartDate &" # And (Calls.Date_Opened)<#" & EndDate &" # )) And (((Calls.Date_Closed)=#09/22/78#)) ;")
	Response.Write "<td>"
	Response.Write " There are "& Count(0) &" Tickets that were Opened or Closed from " & StartDate & " - " & EndDate &"."
	Response.Write " <br>There are "& OpenCount(0) &" Tickets that were Opened from " & StartDate & " - " & EndDate &"."
	Response.Write " <br>There are "& CloseCount(0) &" Tickets that were Closed from " & StartDate & " - " & EndDate &"."
	Response.Write " <br>There are "& SOpenCount(0) &" Tickets that were Opend from " & StartDate & " - " & EndDate &" and are still open."
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
		Response.Write "<td>" & objRs("PROBLEM_DESC") & "</td>"		

		'************************************
		'*  Fill in the Development Hours column, even if there isn't one in the database
		'************************************
		
		objRs.MoveNext
	Wend
	
	
	
	
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

objConn.Close
set objConn=nothing


%>

</BODY>
</HTML>