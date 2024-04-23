<!--#include file="_head.asp"-->
<%
Dim RefPage
RefPage = Request.ServerVariables("PATH_INFO") & "?" & Request.ServerVariables("QUERY_STRING")
'Response.Write RefPage
%>
<script language="Javascript">
	function checknumber(){
		var x=document.TickNum.Num.value
		var anum=/(^\d+$)|(^\d+\.\d+$)/
		if (anum.test(x))
		testresult=true
	else{
		alert("That's not a number. Please Enter a Number.")
		testresult=false
	}
	return (testresult)
}

	function checkban(){
		if (document.layers||document.all)
			return checknumber()
		else
			return true
}
</script>
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
<script language="Javascript">
function print_page(x){
	var y = x
	window.open(y,'print','width=500,height=600,toolbar=0,location=0,directories=0,menuBar=0,scrollBars=1,resizable=1,status=yes' )
}
function uploadfile(){
win3=window.open("uploads/index_mod.asp?TicketNum=<%=Request("Num")%>&RefPage=<%Response.Write RefPage%>","win3","width=375,height=95,titlebar=0,scrollbars=0")
win3.creator=self
}
</script>
<font size=2 face="Arial">
<%
Dim sqlId
Dim sqlId2
Dim objConn
Dim objRec
Dim rsId
Dim rsName
Dim objCnn
Dim objRs
Dim CallWorker
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
sqlId ="SELECT * FROM Calls WHERE Ticket_Number=" & TickNum & ";"

objRec.Open sqlId, objConn	

If objRec.EOF Then
	Response.Write "<font color=red><B>Please enter a valid ticket number!</B></font>"
	Response.Write "<form action=viewticket.asp onSubmit='return checkban()' name='TickNum'>"
	Response.Write "<input type='text' name='Num' size='20'><BR>"
	Response.Write "<input type=Submit value='Get Ticket'>"
Else
	
	If objRec("CLOSED") = "Yes" Then
		Response.Write "<font color=red><B>This ticket has been closed!</B></font><P>"
	End If
	If objRec("CLOSED") = "No" Then
		Response.Write "<font color=green><B>This ticket has not been completed!</B></font><P>"
	End If
	
	Response.Write "<font>"
	Response.Write "<table border=0><tr>"
	Response.Write "<td><P><B>Ticket #:</B> " & objRec("Ticket_Number")&"</td>"
	Response.Write "<td><P>&nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;</td>"
	Response.Write "<td><P><B>Date Opened:</B> " & objRec("Date_Opened")&"</td></tr>"
	Response.Write "<tr><td><B>Call Type:</B> " & objRec("Call_Type") & "</td>"
	Response.Write "<td><P>&nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;</td>"
	Response.Write "<td><P><B>Time Opened:</B> " & objRec("Time_Opened")&"</td></tr>"
	Response.Write "<TR><td><B>Opened By:</B> " & objRec("OpenedBy") & "</td></tr></table>"
	Response.Write "<Table border=0 width='50%'><tr><td><P><B>User Name:</B> " & objRec("User_First_Name") &" "& objRec("User_Last_Name") &"</td></tr>"
	
	Set objConn1 = Server.CreateObject ("ADODB.Connection")
	objConn1.Open "Helpdesk2"
	Country=objRec("country")
	Set objREC1=objConn1.Execute("Select * from Domains WHERE Country = '" & Country & "'")
	Response.Write "<tr><td><B>Email: </B>" & objRec("Email")& objRec1("Domain") & "</TD></tr>"
	Response.Write "<TR><TD><B>Contact&nbsp;Number:&nbsp;</B>" & objRec("USER_CONTACT_NUMBER") &" </td></tr>"
	objRec1.Close
	objConn1.Close
	Set objRec1 = Nothing
	Set objConn1 = Nothing	
	If session("Email") = objRec("Email") and session("Country") = objRec("Country") then
		Response.Write "<form method=post action=ViewTicketSave.asp?Num=" & Request("Num") & ">"
		Response.Write "<tr><td colspan=2><P><B>Problem Description:</B> (Read only)<BR><textarea Readonly cols=60 rows=8 name=Problem_Desc style=FONT-FAMILY: Arial, Geneva, Helvetica, Helv>" & objRec("Problem_Desc") &"</textarea><br><br><b>Additional Info: </b>(Please provide additional information below.)<BR><textarea  cols=60 rows=4 name=Problem_Desc2 style=FONT-FAMILY: Arial, Geneva, Helvetica, Helv></textarea><p></td></tr></table>"
		'Response.Write "<Table border=1><tr><td><P><B>ROI: </B><br><textarea cols=60 rows=2 name=ROI style=FONT-FAMILY: Arial, Geneva, Helvetica, Helv>" & objRec("ROI") &"</textarea></td></tr></table>"		
		Response.Write "<HR width='50%' align=left size=2 noshade color=black>"
		Response.Write "<table border=0>"
		Response.Write "<tr><td><B>In Process?:</B></td><td>" 
		IF objRec("In_Process") ="Yes" Then
		 	Response.Write "In Process<BR></td></tr>"
		End if		
		IF objRec("In_Process") ="MayBe" Then
		 	Response.Write "Pending<BR></td></tr>"
		End if		
		IF objRec("In_Process") ="No" Then
		 	Response.Write "Stopped<BR></td></tr>"
		End if		
		Response.Write"<tr><td><p><B>Package Name:</B></td><td>" & objRec("Package_Name") & "</td></tr>"
		'Response.Write objRec("Call_Serviced_By")
		CallWorker = objRec("Call_Serviced_By")
	
		Set objCnn = Server.CreateObject ("ADODB.Connection")
		Set objRs = Server.CreateObject ("ADODB.Recordset")
	
		objCnn.Open "HelpDesk2"
	
		sqlId2 ="SELECT * FROM Workers WHERE Worker_ID='" & CallWorker & "';"
	
		objRs.Open sqlId2, objCnn

		Response.Write"<TR><td><p><B>Person Working Call:</B></td><td>"
		If CallWorker <> "" Then
			Response.Write "<input type=Hidden name='Call_Serviced_By'  value ="& CallWorker &">" & objRs("First_Name") & " " & objRs("Last_Name")
		End If
		Response.Write "</td></tr>"
		Response.Write"<TR><td><p><B>Expected Production Transport Date:</B></td><td>"
		Response.Write ""& objRec("TransportDate") &"</td></tr>"
		Response.Write "</table>"	

		Response.Write "<table border=0 width='50%'>"
		Response.Write "<tr><td colspan=2><P><B>Solution Description:</B> <br><textarea cols=60 rows=15 name=Solution_Desc  style=FONT-FAMILY: Arial, Geneva, Helvetica, Helv>" & objRec("Solution_Desc") &" </textarea></td></tr></table>"
	
		Response.Write "<table border=0>"
		If objRec("Closed") = "Yes" Then
		Response.Write "<table Border=1><tr><td><font size=1 face=Arial>Closed?:</td><td>&nbsp;<font size=1 face=Arial>Yes:<input type=radio value = Yes name=Closed checked> No:<input type=radio value=No name=Closed><BR></td></tr>"
	End If
	If objRec("Closed") = "No" Then
		Response.Write "<table Border=1><tr><td><font size=1 face=Arial>Closed?:</td><td>&nbsp;<font size=1 face=Arial>Yes:<input type=radio value = Yes name=Closed onMouseDown='this.form.Date_Closed.value=showdate(this)' onMouseUp='this.form.Time_Closed.value=gettime(this)'> No:<input type=radio value=No name=Closed checked onMouseDown='this.form.Date_Closed.value=cleartext()' onMouseUp='this.form.Time_Closed.value=cleartext()'><BR></td></tr>"
	End If

	If objRec("Date_Closed")="9/22/78" Then
	Response.Write"<tr><td><font size=1 face=Arial><p>Date Closed:</td><td><font size=1 face=Arial><input type=text name=Date_Closed></td>"	
	Response.Write"<input type=Hidden name=Date_Closed2  value ='09/22/78'>"
	Else
	Response.Write"<tr><td><font size=1 face=Arial><p>Date Closed:</td><td><font size=1 face=Arial><input type=text name=Date_Closed value = "& objRec("Date_Closed") & "></td>"	
	End if
	If objRec("Time_Closed")="9/22/78" Then
	Response.Write"<td><font size=1 face=Arial><p>Time Closed:</td><td><font size=1 face=Arial><input type=text name=Time_Closed  value = "& objRec("Time_Closed") & "></td></tr>"	
	Response.Write"<input type=Hidden name=Time_Closed2  value ='09/22/78'>"
	Else
	Response.Write"<td><font size=1 face=Arial><p>Time Closed:</td><td><font size=1 face=Arial><input type=text name=Time_Closed  value = "& objRec("Time_Closed") & "></td></tr>"	
	End if
	If objRec("File_Upload_Location") <> "" Then
		Response.Write "<TD><font size=1 face=Arial>File Attachment: </td><td><font size=1 face=Arial>&nbsp;<a href='" & objRec("File_Upload_Location") & "' target='new'>Yes</a>"
	Else
		Response.Write "<TD><font size=1 face=Arial>File Attachment: </td><td><font size=1 face=Arial>No"
	End If
	Response.Write "<input type=button onClick=uploadfile() value='Attach File'>"
	Response.Write "<tr><td colspan=2><input type=Submit  value=Update Information  style='FONT-FAMILY: Arial, Geneva, Helvetica, Helv'></td></form>"
	Else
		Response.Write "<tr><td colspan=2><P><B>Problem Description:</B><BR><font>" & objRec("Problem_Desc") & "</font><p></td></tr></table>"
		Response.Write "<HR width='50%' align=left size=2 noshade color=black>"
		Response.Write "<table border=0>"
		Response.Write "<tr><td><B>In Process?:</B></td><td>" 
		IF objRec("In_Process") ="Yes" Then
		 	Response.Write "In Process<BR></td></tr>"
		End if		
		IF objRec("In_Process") ="MayBe" Then
		 	Response.Write "Pending<BR></td></tr>"
		End if		
		IF objRec("In_Process") ="No" Then
		 	Response.Write "Stopped<BR></td></tr>"
		End if		
		Response.Write"<tr><td><p><B>Package Name:</B></td><td>" & objRec("Package_Name") & "</td></tr>"
		'Response.Write objRec("Call_Serviced_By")
		CallWorker = objRec("Call_Serviced_By")
	
		Set objCnn = Server.CreateObject ("ADODB.Connection")
		Set objRs = Server.CreateObject ("ADODB.Recordset")
	
		objCnn.Open "HelpDesk2"
	
		sqlId2 ="SELECT * FROM Workers WHERE Worker_ID='" & CallWorker & "';"
	
		objRs.Open sqlId2, objCnn

		Response.Write"<TR><td><p><B>Person Working Call:</B></td><td>"
		If CallWorker <> "" Then
			Response.Write objRs("First_Name") & " " & objRs("Last_Name")
		End If
		Response.Write "</td></tr>"
		Response.Write "</table>"	

		Response.Write "<table border=0 width='50%'>"
		Response.Write "<tr><td colspan=2><P><B>Solution Description:</B> <br>" & objRec("Solution_Desc") & "<p></td></tr></table>"
	
		Response.Write "<table border=0>"
		If objRec("Date_Closed") ="9/22/78" then
			Response.Write"<tr><td><p><B>Date Closed: </B></td><td></td></tr>"	
		else
			Response.Write"<tr><td><p><B>Date Closed: </B></td><td>" & objRec("Date_Closed") & "</td></tr>"	
		End If
		If objRec("Time_Closed") ="9/22/78" then
			Response.Write"<TR><td><p><B>Time Closed: </B></td><td></td></tr>"
		else
			Response.Write"<TR><td><p><B>Time Closed: </B></td><td>" & objRec("Time_Closed") & "</td></tr>"	
		End If
		
	End if
	Response.Write "<P></TD></TR></TABLE>"

	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	objRec.Close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing
End If	
%>