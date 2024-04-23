<%Pagename="ShowTicket"%>
<!--#include file="_head.asp"-->
<%
Dim RefPage
RefPage = Request.ServerVariables("PATH_INFO") & "?" & Request.ServerVariables("QUERY_STRING")
'Response.Write RefPage
%>

<%Session("URL") = ""%>
<script language="Javascript">
function showdate(){
	var mydate=new Date()
	var year=mydate.getYear()

	if (year < 1000)
	year+=1900
	var day=mydate.getDay()
	var month=mydate.getMonth()+1
	if (month<10)
	month="0"+month
	var daym=mydate.getDate()
	if (daym<10)
	daym="0"+daym
	var thedate=month+"/"+daym+"/"+year
	return thedate;

}
function gettime(){
	var Digital=new Date()
	var hours=Digital.getHours()
	var minutes=Digital.getMinutes()
	var seconds=Digital.getSeconds()
	var dn="AM" 
	
	if (hours>12){
		dn="PM"
		hours=hours-12
	}
	if (hours==0)
		hours=12
	if (minutes<=9)
		minutes="0"+minutes
	if (seconds<=9)
		seconds="0"+seconds
	
	var thetime=hours+":"+minutes+":"+seconds+" "+dn
	return thetime;
}

function cleartext(){
	var strText = ""
	return strText
}

</script>


<%
	Dim sqlId
	Dim objConn
	Dim objRec
	Dim rsId
	Dim rsName
	Dim CallWorker
	Dim CFDate 
	Dim TransportDate
Dim TickNum
Response.Write "<td valign=top>"
If Request("Num") = "" Then
	TickNum = 0
Else
	TickNum = Request("Num")
End If

Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
	
objConn.Open "HelpDesk2"

Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
	
objConn.Open "HelpDesk2"
sqlId ="SELECT * FROM Calls WHERE Ticket_Number=" & TickNum & ";"

objRec.Open sqlId, objConn	
If objRec.EOF Then
	Response.Write "<font color=red><B>Please enter a valid ticket number!</B></font>"
	Response.Write "<form action=Modifyticket.asp onSubmit='return checkban()' name='TickNum'>"
	Response.Write "<input type='text' name='Num' size='20'><BR>"
	Response.Write "<input type=Submit value='Get Ticket'>"
	Response.Write "</form>"
Else


	
	Response.Write "<table border=1 cellpadding=2><tr>"
	Response.Write "<td><P>Ticket #: " & objRec("Ticket_Number")&"</td>"
	
	Response.Write "<td><P>Date Opened: " & objRec("Date_Opened")&"</td></tr>"
		Response.Write "<P><form method=post action=SaveCallBack.asp?Num=" & Request("Num") & " onsubmit='return Join_Form1_Validator(this)'>"
	Response.Write "<tr><td>Call Type: " & objRec("Call_Type") & "</td>"
	Response.Write "<td><P>Time Opened: " & objRec("Time_Opened")&"</td></tr>"
	Response.Write "<TR><td colspan=2>Opened By: " & objRec("OpenedBy") & "</td></tr>"
	Response.Write "<tr><td colspan=2><P>User's Name: " & objRec("User_First_Name") &" "& objRec("User_Last_Name") &"</td>"
	Response.Write "</tr>"
	
	Set objConn1 = Server.CreateObject ("ADODB.Connection")
	objConn1.Open "Helpdesk2"
	Country=objRec("country")
	Email = objRec("Email")
	Set objREC1=objConn1.Execute("Select * from Domains WHERE Country = '" & Country & "'")
	Response.Write "<tr><td>E-mail: " & objRec("Email") & objRec1("Domain") &"</td><td>Phone: " & objRec("USER_CONTACT_NUMBER") &"</td></tr></table>"
	Response.Write "<input type=Hidden Name=Email Value=" & objRec("Email")& objRec1("Domain")&"><input type=Hidden Name=UFName Value=" & objRec("User_First_Name") &"><input type=Hidden Name=ULName Value=" & objRec("User_Last_Name") &">"
	Response.Write "<Table border=1><tr><td><P>Problem Description: (Read only)<br><textarea Readonly cols=47 rows=8 name=Problem_Desc style=FONT-FAMILY: Arial, Geneva, Helvetica, Helv>" & objRec("Problem_Desc") &"</textarea></td></tr></table>"
	objRec1.Close
	objConn1.Close
	Set objRec1 = Nothing
	Set objConn1 = Nothing	

	Response.Write "<Table border=1><tr><td colspan=2><P>Solution Description: (Read only)<br><textarea Readonly cols=47 rows=8 name=Solution_Desc  style=FONT-FAMILY: Arial, Geneva, Helvetica, Helv>" & objRec("Solution_Desc") &" </textarea></td></tr>"
	Response.Write"<tr><td><p>Date Closed: "& objRec("Date_Closed") & "</td>"	
	Response.Write"<td><p>Time Closed: "& objRec("Time_Closed") & "</td></tr>"
	Response.Write "<Table border=1><tr><td colspan=4><P>Comments: <br><textarea cols=47 rows=8 name=Comments  style=FONT-FAMILY: Arial, Geneva, Helvetica, Helv>" & objRec("Call_Back_Comments") &" </textarea></td></tr>"	
	Response.Write "<td colspan=2><font size=1 face=Arial>Call Back Complete:<input type=CheckBox name=Call_Back_Complete unchecked></td>"
	Response.Write "<td colspan=1><input type=submit value='Update Information' style='FONT-FAMILY: Arial, Geneva, Helvetica, Helv'></td></tr></table></form></table>"


End if

	objRec.Close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing

%>

</TD>
</TR>
</TABLE>
</BODY>
</HTML>

