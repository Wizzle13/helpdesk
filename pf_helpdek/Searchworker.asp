<!--#include file="_head.asp"-->

<td>
<%

Dim rsCount
Dim LCount
Dim pCount
Dim sCount
Dim strView
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"
RsCount = 0
If Request("View") ="All" Then
	if Request("search") ="Worker" then
		Set objRs=objConn.Execute("Select * from Calls Where Call_Serviced_By =  '"& Request.Form("ServicedBy") &"' order by Ticket_Number")
		Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE CALL_SERVICED_BY =  '"& Request.Form("ServicedBy") &"'")
	End if

	if Request("search") ="Email" then
		Logon=Request.Form("Logon_Name")
		Country=Request.Form("Domain")
		Set objRs=objConn.Execute("Select * from Calls WHERE Email = '" & Logon & "' and Country='" & Country & "' order by Ticket_Number")
		Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Email = '" & Logon & "'")
	End if

	if Request("search") ="FName" then
		Set objRs=objConn.Execute("Select * from Calls WHERE User_First_Name = '" & Request.Form("FirstName") & "' order by Ticket_Number")
		Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE User_First_Name = '" & Request.Form("FirstName") & "'")
	End if

	if Request("search") ="LName" then
		Set objRs=objConn.Execute("Select * from Calls WHERE User_Last_Name = '" & Request.Form("LastName") & "' order by Ticket_Number")
		Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE User_Last_Name = '" & Request.Form("LastName") & "'")
	End if
	
	if Request("search") ="SAP_Module" then
		Set objRs=objConn.Execute("Select * from Calls WHERE SAP_Module =  '"& Request.Form("SAP_Module") &"' order by Ticket_Number")
		Set rsCount=objConn.Execute("Select Count(*) from Calls WHERE SAP_Module =  '"& Request.Form("SAP_Module") &"'")	
	End if
	
	if Request("search") ="PackageName" then
		Set objRs=objConn.Execute("Select * from Calls WHERE Package_name =  '"& Request.Form("PackageName") &"' order by Ticket_Number")
		Set rsCount=objConn.Execute("Select Count(*) from Calls WHERE Package_name =  '"& Request.Form("PackageName") &"'")	
	End if

Else

	if Request("search") ="Worker" then
		Set objRs=objConn.Execute("Select * from Calls WHERE Closed = '" & Request("View") & "' and  Call_Serviced_By =  '"& Request.Form("ServicedBy") &"' order by Ticket_Number")
		Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Closed = '" & Request("View") & "' and CALL_SERVICED_BY =  '"& Request.Form("ServicedBy") &"'")
	End if

	if Request("search") ="Email" then
		Logon=Request.Form("Logon_Name")
		Country=Request.Form("Domain")
		Set objRs=objConn.Execute("Select * from Calls WHERE Closed = '" & Request("View") & "' and Email = '" & Logon & "' order by Ticket_Number")
		Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Closed = '" & Request("View") & "' and Email = '" & Logon & "'")
	End if	
	
	if Request("search") ="FName" then
		Set objRs=objConn.Execute("Select * from Calls WHERE Closed = '" & Request("View") & "' and User_First_Name = '" & Request.Form("FirstName") & "' order by Ticket_Number")
		Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Closed = '" & Request("View") & "' and User_First_Name = '" & Request.Form("FirstName") & "'")
	End if

	if Request("search") ="LName" then
		Set objRs=objConn.Execute("Select * from Calls WHERE Closed = '" & Request("View") & "' and User_Last_Name = '" & Request.Form("LastName") & "' order by Ticket_Number")
		Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Closed = '" & Request("View") & "' and User_Last_Name = '" & Request.Form("LastName") & "'")
	End if
	
	if Request("search") ="SAP_Module" then
		Set objRs=objConn.Execute("Select * from Calls WHERE Closed = '" & Request("View") & "' and SAP_Module = '" & Request.Form("SAP_Module") & "' order by Ticket_Number")
		Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Closed = '" & Request("View") & "' and SAP_Module = '" & Request.Form("SAP_Module") & "'")
	End if
	
	if Request("search") ="PackageName" then
		Set objRs=objConn.Execute("Select * from Calls WHERE Closed = '" & Request("View") & "' and Package_name =  '"& Request.Form("PackageName") &"' order by Ticket_Number")
		Set rsCount=objConn.Execute("Select Count(*) from Calls WHERE Closed = '" & Request("View") & "' and Package_name =  '"& Request.Form("PackageName") &"'")
	End if
	
End IF
Response.Write "<a href=""search.asp"">Search</a>"
Response.Write "<form method=post action=ExportSearchWorker.asp target=New>"
If Request("search") ="Worker" then
	Response.Write "<input Type='hidden' Name='Search' Value='Worker'>"
	Response.Write "<input Type='hidden' Name='ServicedBy' Value='"& Request.Form("ServicedBy") &"'>"
End if
If Request("search") ="FName" then
	Response.Write "<input Type='hidden' Name='Search' Value='FName'>"
	Response.Write "<input Type='hidden' Name='FirstName' Value='"& Request.Form("FirstName") &"'>"
End if
If Request("search") ="LName" then
	Response.Write "<input Type='hidden' Name='Search' Value='LName'>"
	Response.Write "<input Type='hidden' Name='LastName' Value='"& Request.Form("LastName") &"'>"
End if
If Request("search") ="Email" then
	Response.Write "<input Type='hidden' Name='Search' Value='Email'>"
	Response.Write "<input Type='hidden' Name='Logon_Name' Value='"& Request.Form("Logon_Name") &"'>"
	Response.Write "<input Type='hidden' Name='Domain' Value='"& Request.Form("Domain") &"'>"
End if
If Request("search") ="SAP_Module" then
	Response.Write "<input Type='hidden' Name='Search' Value='SAP_Module'>"
	Response.Write "<input Type='hidden' Name='SAP_Module' Value='"& Request.Form("SAP_Module") &"'>"
End if
If Request("search") ="PackageName" then
	Response.Write "<input Type='hidden' Name='Search' Value='PackageName'>"
	Response.Write "<input Type='hidden' Name='PackageName' Value='"& Request.Form("PackageName") &"'>"
End if
Response.Write "<input Type='Hidden' Name='View' Value='"& Request.Form("View") &"'>"
Response.Write "<input type=Submit value='Export to Excel'></form>"
If rsCount(0) <> 0 Then
	Response.Write "<br>There are "& RsCount(0) &" Tickets.<br>"
	Response.Write "<table border=0><TR><TD bgcolor=#A8EAA8>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Working on</TD><TD>&nbsp;&nbsp;</TD>"
	Response.Write "<TD bgcolor=#FBFD04>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Pending</TD><TD>&nbsp;&nbsp;</TD>"
	Response.Write "<TD bgcolor=#FF6666>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Stopped</TD></TR></TABLE>"
	Response.Write "<table border=1 cellpadding=0 cellborder=0>"
	Response.Write "<tr><td><B>Ticket #</td>"
	Response.Write "<td><B>User<br>Name</td>"
	Response.Write "<td><B>SAP<br>Priority</td>"
	Response.Write "<td><B>SAP<br>Module</td>"
	Response.Write "<td><B>Call<br>Type</td>"
	Response.Write "<td><B>Date<BR>Opened</td>"
	Response.Write "<td><B>Projected<br>Date<BR>Closed</td>"
	Response.Write "<td><B>Date<BR>Closed</td>"
	Response.Write "<td><B>Serviced<br>By</td>"
	Response.Write "<td width= 50><B>Problem</td>"
	Response.Write "<td width= 50><B>ROI</td>"
	Response.Write "<td width= 50><B>Solution</td>"

	While Not objRs.EOF
		If objRs("Closed") = "Yes" Then
			Response.Write "<tr bgcolor=lightgrey>"
		Else
			Response.Write "<tr>"
		End If

		Response.Write "<td align=center><a href=""modifyticket.asp?Num=" & objRs("Ticket_Number") & """>" & objRs("Ticket_Number") & "</a>" & "</td>"
		Response.Write "<td>" & objRs("User_First_Name") &"<br>"& objRs("User_Last_Name") & "</td>"	
		Response.Write "<td><center>" & objRs("SAPPriority") & "</center></td>"
		Response.Write "<td><center>" & objRs("SAP_Module") & "</center></td>"
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
		Response.Write "<td>" & objRs("Projected_Complete_Date") & "</td>"
		IF objRs("Date_Closed")="9/22/78" then
			Response.Write "<td></td>"
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
		ROI = objRs("ROI")
		If Len(ROI) > 100 Then						
			Response.Write "<td>" & Mid(ROI, 1, 100) & "...<a href=""modifyticket.asp?Num=" & objRs("Ticket_Number") & """>(more)</a></td>"		
		Else
			Response.Write "<td>" & ROI & "</td>"
		End If
		SOLUTION_DESC = objRs("SOLUTION_DESC")
		If Len(SOLUTION_DESC) > 100 Then						
			Response.Write "<td>" & Mid(SOLUTION_DESC, 1, 100) & "...<a href=""modifyticket.asp?Num=" & objRs("Ticket_Number") & """>(more)</a></td>"		
		Else
			Response.Write "<td>" & SOLUTION_DESC & "</td>"
		End If
		'************************************
		'*  Fill in the Development Hours column, even if there isn't one in the database
		'************************************
		
		objRs.MoveNext
	Wend
Else
	Response.Write "<BR>There are no tickets that fit your selection.<P>"
End If
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing%>

</BODY>
</HTML>