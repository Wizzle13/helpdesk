<html>
<head>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
var scroller_array = new Array();
function start_scroller(){
	for (var i = 0; i < scroller_array.length; i++) {
		eval(scroller_array[i]);
	}
}
function common_onload(){
  start_scroller();
  
  		window.print()
		window.close();
  
}
</script>	


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
</head>
<body onLoad="common_onload()">

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
	
	
	
	Response.Write "<font>"
	Response.Write "<table border=0><tr>"
	Response.Write "<td><P><B>Ticket #:</B> " & objRec("Ticket_Number")&"</td>"
	Response.Write "<td><P>&nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;</td>"
	Response.Write "<td><P><B>Date Opened:</B> " & objRec("Date_Opened")&"</td></tr>"
	Response.Write "<tr><td><B>Call Type:</B> " & objRec("Call_Type") & "</td>"
	Response.Write "<td><P>&nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;</td>"
	Response.Write "<td><P><B>Time Opened:</B> " & objRec("Time_Opened")&"</td></tr>"
	Response.Write "<TR><td><B>Opened By:</B>" & objRec("OpenedBy") & "</td></tr></table>"

	Response.Write "<Table border=0 width='50%'><tr><td><P><B>User Name:</B> " & objRec("User_First_Name") &" "& objRec("User_Last_Name") &"</td></tr>"
	
	Set objConn1 = Server.CreateObject ("ADODB.Connection")
	objConn1.Open "Helpdesk2"
	Country=objRec("country")
	Set objREC1=objConn1.Execute("Select * from Domains WHERE Country = '" & Country & "'")
	Response.Write "<tr><td>>Email: " & objRec("Email")& objRec1("Domain") & "</TD></tr>"
	Response.Write "<TR><TD><B>Contact&nbsp;Number:&nbsp;</B>" & objRec("USER_CONTACT_NUMBER") &" </td></tr>"
	objRec1.Close
	objConn1.Close
	Set objRec1 = Nothing
	Set objConn1 = Nothing	
	
	Response.Write "<tr><td colspan=2><P><B>Problem Description:</B><BR><font>"& objRec("PROBLEM_DESC")& "</font><p></td></tr>"
	Response.Write "<tr><td colspan=2><P><B>ROI:</B><BR><font>"& objRec("ROI")& "</font><p></td></tr></table>"
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
	Response.Write "<tr><td><B>Priority:</B></td><td>"
	IF objRec("Priority") ="1" Then
		 Response.Write "High<BR></td></tr>" 	
	End if			 
	IF objRec("Priority") ="2" Then
		 Response.Write "Normal<BR></td></tr>" 	
	End if			 
	IF objRec("Priority") ="3" Then
		 Response.Write "Low<BR></td></tr>" 	
	End if
	Response.Write "<tr><td><B>SAP Priority:</B></td><td>" & objRec("SAPPriority")& "</td></tr>"
	Response.Write "<tr><td><B>SAP Upgrade:</B></td><td>"
	If objRec("Upgrade") = "on" then
		Response.Write "Yes<BR></td></tr>" 
	Else 
		Response.Write "No<BR></td></tr>" 
	End if
	Response.Write "<tr><td><B>SAP Module:</B></td><td>" & objRec("SAP_Module")& "</td></tr>"	
	Response.Write "<tr><td><B>Expected Production Transport Date:</B></td><td>" & objRec("TransportDate")& "</td></tr>"				 
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
	Response.Write"<tr><td><p><B>Projected Complete Date: </B></td><td>" & objRec("Projected_Complete_Date") & "</td></tr>"	
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