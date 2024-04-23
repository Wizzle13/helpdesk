<!--#include file="_head.asp"-->
<td>
<%

Dim rsCount
Dim LCount
Dim pCount
Dim sCount
Dim strView
Dim strCountry
Dim strSearch

RsCount = 0

strCountry = Request("Country")
strCountry = Replace( strCountry, ", ", "','")

strView = Request.Form("CountryView")
If strView = "Open" Then
	strSearch = " AND Closed ='No'"
End If
If strView = "Closed" Then
	strSearch = " AND Closed ='Yes'"
End If
If strView = "" Then
	strSearch = ""
End If
	

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"

Set objRs=objConn.Execute("Select * From Calls WHERE Country IN ('"& strCountry &"')" & strSearch & " order by Ticket_Number")
Set rsCount=objConn.Execute("SELECT Count(*) from Calls WHERE Country IN ('"& strCountry &"')" & strSearch & "")

Response.Write "<a href=""search.asp"">Search Again</a><BR><P>"
'Response.Write strView & vbcrlf & strSearch



If rsCount(0) <> 0 Then
	Response.Write "<br>There are "& RsCount(0) &" tickets that match your country selection.<br>"
	Response.Write "<table border=0><TR><TD bgcolor=#A8EAA8>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Working on</TD><TD>&nbsp;&nbsp;</TD>"
	Response.Write "<TD bgcolor=#FBFD04>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Pending</TD><TD>&nbsp;&nbsp;</TD>"
	Response.Write "<TD bgcolor=#FF6666>&nbsp;&nbsp;&nbsp;</TD><TD><font color=black><B>Stopped</TD></TR></TABLE>"
	Response.Write "<table border=1 cellpadding=0 cellborder=0>"
	Response.Write "<tr><td><B>Ticket #</td>"
	Response.Write "<td><B>User<br>Name</td>"
	Response.Write "<td><B>Call<br>Type</td>"
	Response.Write "<td><B>Date<BR>Opened</td>"
	Response.Write "<td><B>Serviced<br>By</td>"
	Response.Write "<td><B>Country</td>"
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
		Response.Write "<td>" & objRs("CALL_SERVICED_BY") & "</td>"
		Response.Write "<td>" & objRs("Country") & "</td>"
		ProblemDesc = objRs("Problem_Desc")						
		Response.Write "<td>" & Mid(ProblemDesc, 1, 100) &Mid(ProblemDesc, 1, 100) & "...<a href=""modifyticket.asp?Num=" & objRs("Ticket_Number") & """>(more)</a></td>"		

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