<%Pagename="ShowTicket"%>
<!--#include file="_head.asp"-->
<%
Dim RefPage
RefPage = Request.ServerVariables("PATH_INFO") & "?" & Request.ServerVariables("QUERY_STRING")
'Response.Write RefPage
%>

<%Session("URL") = ""%>
<script language="Javascript">
function checknumber()
{
		var x=document.TickNum.Num.value
		var anum=/(^\d+$)|(^\d+\.\d+$)/
		if (anum.test(x))
		testresult=true
	else{
		alert("That's not a number.  Please enter a number.")
		testresult=false
	}
	return (testresult)
}

function checkban()
{
		if (document.layers||document.all)
			return checknumber()
		else
			return true
}

function Join_Form1_Validator(theForm)
{

  if (theForm.Package_Name.value == "")
  {
    alert("Please enter a value for the \"Package Name\" field.");
    theForm.Package_Name.focus();
    return (false);
  }
  if (theForm.Package_Name.value == "_Unclassified_")
  {
    alert("Please enter a value for the \"Package Name\" field.");
    theForm.Package_Name.focus();
    return (false);
  }
return (true);
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


	Response.Write "<font face=Arial>"
	Response.Write "<P><a href=delete.asp?Num=" & objRec("Ticket_Number") & ">Delete</a>&nbsp;&nbsp; | "
	Response.Write "&nbsp;<a href=javascript:print_page('viewonly.asp?Num=" & objRec("Ticket_Number") & "')>Print This Ticket</a><P>"
	
	Response.Write "<table border=1 cellpadding=2><tr>"
	Response.Write "<td><P>Ticket #: " & objRec("Ticket_Number")&"</td>"
	
	Response.Write "<td><P>Date Opened: " & objRec("Date_Opened")&"</td></tr>"
		Response.Write "<P><form method=post action=saveticket.asp?Num=" & Request("Num") & " onsubmit='return Join_Form1_Validator(this)'>"
	If objRec("Call_Type") = "HELP" Then
		Response.Write "<tr><td>Call Type: <select name=Call_Type style='FONT-FAMILY: Arial, Geneva, Helvetica, Helv'><option selected>HELP<option>REQUEST<option>HELP-3rd Party</select><br></td>"
	End If
	If objRec("Call_Type") = "REQUEST" Then
		Response.Write "<tr><td>Call Type: <select name=Call_Type style='FONT-FAMILY: Arial, Geneva, Helvetica, Helv'><option>HELP<option selected>REQUEST<option>HELP-3rd Party</select><br></td>"
	End If
	If objRec("Call_Type") = "HELP-3rd Party" Then
		Response.Write "<tr><td>Call Type: <select name=Call_Type style='FONT-FAMILY: Arial, Geneva, Helvetica, Helv'><option>HELP<option>REQUEST<option selected>HELP-3rd Party</select><br></td>"
	End If
	
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
	Response.Write "<Table border=1><tr><td><P>Problem Description: <br><textarea cols=60 rows=8 name=Problem_Desc style=FONT-FAMILY: Arial, Geneva, Helvetica, Helv>" & objRec("Problem_Desc") &"</textarea></td></tr></table>"
	Response.Write "<Table border=1><tr><td><P>ROI: <br><textarea cols=60 rows=8 name=ROI style=FONT-FAMILY: Arial, Geneva, Helvetica, Helv>" & objRec("ROI") &"</textarea></td></tr></table>"
	objRec1.Close
	objConn1.Close
	Set objRec1 = Nothing
	Set objConn1 = Nothing	
	Response.Write "<HR width=600 align=left size=5 noshade color=black>"

	'***************************************************
	'*  Set the value for the Priority radio buttons according to the value in the database
	'***************************************************
	'***************************************************
	'*  Set the value for the Completed radio buttons according to the value in the database
	'***************************************************
	Response.Write "<table border=1>"
	If objRec("IN_PROCESS") = "Yes" Then
		Response.Write "<tr><td><font size=1 face=Arial>Status:</td><td><font size=1 face=Arial><input type=radio value =Yes name=IN_PROCESS checked>In Process<BR><input type=radio value=MayBe name=IN_PROCESS>Pending<BR><input type=radio value=No name=IN_PROCESS>Stopped</td>"
	End If
	If objRec("IN_PROCESS") = "MayBe" Then
		Response.Write "<tr><td><font size=1 face=Arial>Status:</td><td><font size=1 face=Arial><input type=radio value =Yes name=IN_PROCESS>In Process<BR><input type=radio value=MayBe name=IN_PROCESS Checked>Pending<BR><input type=radio value=No name=IN_PROCESS>Stopped</td>"
	End If
	If objRec("IN_PROCESS") = "No" Then
		Response.Write "<tr><td><font size=1 face=Arial>Status:</td><td><font size=1 face=Arial><input type=radio value =Yes name=IN_PROCESS>In Process<BR><input type=radio value=MayBe name=IN_PROCESS>Pending<BR><input type=radio value=No name=IN_PROCESS checked>Stopped</td>"
	End If
	
	If objRec("Priority") = "1" Then
		Response.Write "<td><font size=1 face=Arial>Priority:</td><td><font size=1 face=Arial><input type=radio value =1 name=Priority checked>High<BR><input type=radio value=2 name=Priority>Normal<BR><input type=radio value=3 name=Priority>Low</td></tr>"
	End If
	If objRec("Priority") = "2" Then
		Response.Write "<td><font size=1 face=Arial>Priority:</td><td><font size=1 face=Arial><input type=radio value =1 name=Priority>High<BR><input type=radio value=2 name=Priority Checked>Normal<BR><input type=radio value=3 name=Priority>Low</td></tr>"
	End If
	If objRec("Priority") = "3" Then
		Response.Write "<td><font size=1 face=Arial>Priority:</td><td><font size=1 face=Arial><input type=radio value =1 name=Priority>High<BR><input type=radio value=2 name=Priority>Normal<BR><input type=radio value=3 name=Priority checked>Low</td></tr>"
	End If
	
	Response.Write "<tr><td><font size=1 face=Arial>SAP Priority:</td><td align=center>" & vbcrlf
		Response.Write "<select name=tickpriority>"
		i = 1
		While i <> 21
			If i = objRec("SAPPriority") Then
				Response.Write "<option value=" & objRec("SAPPriority") & " selected>" & objRec("SAPPriority") & vbcrlf
			Else
				Response.Write "<option value=" & i & ">" & i & vbcrlf
			End if
			i = i + 1
		Wend
		If objRec("SAPPriority") = 99 Then
			Response.Write "<option value=" & objRec("SAPPriority") & " selected>" & objRec("SAPPriority") & vbcrlf
		Else
			Response.Write "<option value=99>99" & vbcrlf
		End if
		Response.Write "</select></td>"  & vbcrlf
		strDate=FormatDateTime (Date, vbShortDate)
		Response.Write "<td><font size=1 face=Arial>Expected Production Transport Date:</td>"
		
		IF Session("WorkID")="CA" Then
			Response.Write "<td align=center ><font size=1 face=Arial>"
			Set objRs1=objConn.Execute("Select * from SAP_Transports")
			sqlId ="SELECT * FROM SAP_Transports;"
			Response.Write "<select name=TransportDate>"
			Response.Write "<option value='12/31/2049'>N/A"
			Response.Write "<option Value="& objRec("TransportDate") & " selected>" & objRec("TransportDate")
			
			While Not objRs1.EOF
				If objRs1("Transport_Date") > Date then
					Response.Write"<Option value="& objRs1("Transport_Date")&">"& objRs1("Transport_Date")
				End if
			objRs1.MoveNext
			Wend
			Response.Write "</select></td></tr>"
		Else
			Set objRs1=objConn.Execute("Select * from SAP_Transports")
			sqlId ="SELECT * FROM SAP_Transports;"
			Response.Write "<td align=center ><font size=1 face=Arial>"
				IF objRec("TransportDate")="12/31/2049" or objRec("TransportDate")="''"  Then
						Response.Write"<input type=Hidden name='TransportDate' value='12/31/2049'>N/A"
					Else	
						Response.Write "<input type=Hidden name='TransportDate' value="& objRec("TransportDate")&">"& objRec("TransportDate")
				End if
			Response.Write "</td></tr>"
			
		End IF		
'	If objRec("upgrade") = "on" Then
'		Response.Write "<td><font size=1 face=Arial>SAP Upgrade:</td><td><input type=checkbox name=upgrade checked></td></tr>"
'	Else
'		Response.Write "<td><font size=1 face=Arial>SAP Upgrade:</td><td><input type=checkbox name=upgrade></td></tr>"
'	End if

	Response.Write "<tr><td><font size=1 face=Arial>SAP Module:</td><td align=center>" & vbcrlf
	Set strSAPModule = objRec("SAP_Module")
	'Response.Write "{" & strSAPModule & "}"	

	Response.Write "<select name=SAP_Module>"

	Select Case strSAPModule
		Case ""
			Response.Write "<option value=N/A selected>N/A"
			Response.Write "<option value=SD>SD"
			Response.Write "<option value=MM>MM"
			Response.Write "<option value=PP>PP"
			Response.Write "<option value=FI/CO>FI/CO"
			Response.Write "<option value=Basis>Basis"
			Response.Write"<Option Value=WM>WM"
			Response.Write"<Option Value=BW>BW"
			Response.Write"<Option Value=APO>APO"
		Case "N/A"
			Response.Write "<option value=N/A selected>N/A"
			Response.Write "<option value=SD>SD"
			Response.Write "<option value=MM>MM"
			Response.Write "<option value=PP>PP"
			Response.Write "<option value=FI/CO>FI/CO"
			Response.Write "<option value=Basis>Basis"
			Response.Write"<Option Value=WM>WM"
			Response.Write"<Option Value=BW>BW"
			Response.Write"<Option Value=APO>APO"
		Case "SD"
			Response.Write "<option value=N/A>N/A"
			Response.Write "<option value=SD selected>SD"
			Response.Write "<option value=MM>MM"
			Response.Write "<option value=PP>PP"
			Response.Write "<option value=FI/CO>FI/CO"
			Response.Write "<option value=Basis>Basis"
			Response.Write"<Option Value=WM>WM"
			Response.Write"<Option Value=BW>BW"
			Response.Write"<Option Value=APO>APO"
		Case "MM"
			Response.Write "<option value=N/A>N/A"
			Response.Write "<option value=SD>SD"
			Response.Write "<option value=MM selected>MM"
			Response.Write "<option value=PP>PP"
			Response.Write "<option value=FI/CO>FI/CO"
			Response.Write "<option value=Basis>Basis"
			Response.Write"<Option Value=WM>WM"
			Response.Write"<Option Value=BW>BW"
			Response.Write"<Option Value=APO>APO"
		Case "PP"
			Response.Write "<option value=N/A>N/A"
			Response.Write "<option value=SD>SD"
			Response.Write "<option value=MM>MM"
			Response.Write "<option value=PP selected>PP"
			Response.Write "<option value=FI/CO>FI/CO"
			Response.Write "<option value=Basis>Basis"
			Response.Write"<Option Value=WM>WM"
			Response.Write"<Option Value=BW>BW"
			Response.Write"<Option Value=APO>APO"
		Case "FI/CO"
			Response.Write "<option value=N/A>N/A"
			Response.Write "<option value=SD>SD"
			Response.Write "<option value=MM>MM"
			Response.Write "<option value=PP>PP"
			Response.Write "<option value=FI/CO selected>FI/CO"
			Response.Write "<option value=Basis>Basis"
			Response.Write"<Option Value=WM>WM"
			Response.Write"<Option Value=BW>BW"
			Response.Write"<Option Value=APO>APO"
		Case "Basis"
			Response.Write "<option value=N/A>N/A"
			Response.Write "<option value=SD>SD"
			Response.Write "<option value=MM>MM"
			Response.Write "<option value=PP>PP"
			Response.Write "<option value=FI/CO>FI/CO"
			Response.Write "<option value=Basis selected>Basis"
			Response.Write"<Option Value=WM>WM"
			Response.Write"<Option Value=BW>BW"
			Response.Write"<Option Value=APO>APO"
		Case "WM"
			Response.Write "<option value=N/A>N/A"
			Response.Write "<option value=SD>SD"
			Response.Write "<option value=MM>MM"
			Response.Write "<option value=PP>PP"
			Response.Write "<option value=FI/CO>FI/CO"
			Response.Write "<option value=Basis>Basis"
			Response.Write"<Option Value=WM selected>WM"
			Response.Write"<Option Value=BW>BW"
			Response.Write"<Option Value=APO>APO"
		Case "BW"
			Response.Write "<option value=N/A>N/A"
			Response.Write "<option value=SD>SD"
			Response.Write "<option value=MM>MM"
			Response.Write "<option value=PP>PP"
			Response.Write "<option value=FI/CO>FI/CO"
			Response.Write "<option value=Basis>Basis"
			Response.Write"<Option Value=WM>WM"
			Response.Write"<Option Value=BW selected>BW"
			Response.Write"<Option Value=APO>APO"
		Case "APO"
			Response.Write "<option value=N/A>N/A"
			Response.Write "<option value=SD>SD"
			Response.Write "<option value=MM>MM"
			Response.Write "<option value=PP>PP"
			Response.Write "<option value=FI/CO>FI/CO"
			Response.Write "<option value=Basis>Basis"
			Response.Write"<Option Value=WM>WM"
			Response.Write"<Option Value=BW>BW"
			Response.Write"<Option Value=APO selected>APO"

	End Select

	Response.Write "</select></td></tr>"


	
	Response.Write"<tr><td><font size=1 face=Arial><p>Package Name:</td><td><select name=Package_Name><Option value='"& objRec("Package_name")&"'>" & objRec("Package_Name")
	Set objConn=Server.CreateObject("ADODB.Connection")
	objConn.Open "DSN=HelpDesk2"
	Set objRs=objConn.Execute("Select * from GROUP_MEMBERS order by Package_Name")
	
	While Not objRs.EOF
	Response.Write"<Option value='"& objRs("Package_name")&"'>"& objRs("Package_Name")
		objRs.MoveNext
	Wend
	Response.Write"</select></td>"
	
	objConn.Close
	
	Set objConn = Nothing	
	Set objConn=Server.CreateObject("ADODB.Connection")
	objConn.Open "DSN=HelpDesk2"
	strWorker= objRec("Call_Serviced_By")
	Response.Write "<input type=Hidden Name=strworker Value=" & objRec("Call_Serviced_By") &">"
	Set objRs=objConn.Execute("Select * from Workers where Worker_ID= '"&strWorker&"'")
	Response.Write"<td><font size=1 face=Arial><p>Person Working Call:</td><td><font size=1 face=Arial><select name=Call_Serviced_By>"
	If strWorker<>"" Then
	Response.Write"<Option value= " & objRs("Worker_ID") & ">"& objRs("First_Name")&" "& objRs("Last_Name")
	Else
	Response.Write"<Option>"
	End if
	Set objRs=objConn.Execute("Select * from Workers where Active_Worker='Yes' order by First_Name, Last_Name")
	sqlId ="SELECT * FROM Workers;"
	While Not objRs.EOF
	Response.Write"<Option value="& objRs("Worker_ID")&">"& objRs("First_Name")&" "& objRs("Last_Name")
		objRs.MoveNext
	Wend
	Response.Write"</select></td></tr></table>"

	Response.Write "<Table border=1><tr><td><P>Solution Description: <br><textarea cols=60 rows=15 name=Solution_Desc  style=FONT-FAMILY: Arial, Geneva, Helvetica, Helv>" & objRec("Solution_Desc") &" </textarea></td></tr></table>"

	
	If objRec("Closed") = "Yes" Then
		Response.Write "<table Border=1><tr><td><font size=1 face=Arial>Closed?:</td><td>&nbsp;<font size=1 face=Arial>Yes:<input type=radio value = Yes name=Closed checked> No:<input type=radio value=No name=Closed><BR></td>"
	End If
	If objRec("Closed") = "No" Then
		Response.Write "<table Border=1><tr><td><font size=1 face=Arial>Closed?:</td><td>&nbsp;<font size=1 face=Arial>Yes:<input type=radio value = Yes name=Closed onMouseDown='this.form.Date_Closed.value=showdate(this)' onMouseUp='this.form.Time_Closed.value=gettime(this)'> No:<input type=radio value=No name=Closed checked onMouseDown='this.form.Date_Closed.value=cleartext()' nMouseUp='this.form.Time_Closed.value=cleartext()'><BR></td>"
	End If
	
	If objRec("File_Upload_Location") <> "" Then
		Response.Write "<TD><font size=1 face=Arial>File Attachment: </td><td><font size=1 face=Arial>&nbsp;<a href='" & objRec("File_Upload_Location") & "' target='new'>Yes</a>"
	Else
		Response.Write "<TD><font size=1 face=Arial>File Attachment: </td><td><font size=1 face=Arial>No"
	End If
	
	Response.Write "<input type=button onClick=uploadfile() value='Attach File'>"
	
	Response.Write "</td></tr>"
	Response.Write "<tr><td colspan=3><font size=1 face=Arial><p>Projected Complete Date:</td>"

	If objRec("Projected_Complete_Date")="12/31/2049" Then
		Response.Write "<td><font size=1 face=Arial><input type=Text name=Projected_Complete_Date></td></tr>"
		Response.Write "<input type=Hidden name=Projected_Complete_Date2 value='12/31/2049'>"
	Else 
		Response.Write "<td><font size=1 face=Arial><input type=text name=Projected_Complete_Date value = "& objRec("Projected_Complete_Date") & "></td></tr>"
		Response.Write "<input type=Hidden name=Projected_Complete_Date2 value='12/31/2049'>"
	End If
	If objRec("Date_Closed")="9/22/78" Then
	Response.Write"<tr><td><font size=1 face=Arial><p>Date Closed:</td><td><font size=1 face=Arial><input type=text name=Date_Closed></td>"	
	Response.Write"<input type=Hidden name=Date_Closed2  value ='09/22/78'>"
	Else
	Response.Write"<tr><td><font size=1 face=Arial><p>Date Closed:</td><td><font size=1 face=Arial><input type=text name=Date_Closed value = "& objRec("Date_Closed") & "></td>"	
	Response.Write"<input type=Hidden name=Date_Closed2  value ='09/22/78'>"
	End if
	If objRec("Time_Closed")="9/22/78" Then
	Response.Write"<td><font size=1 face=Arial><p>Time Closed:</td><td><font size=1 face=Arial><input type=text name=Time_Closed></td></tr>"	
	Response.Write"<input type=Hidden name=Time_Closed2  value ='09/22/78'>"
	Else
	Response.Write"<td><font size=1 face=Arial><p>Time Closed:</td><td><font size=1 face=Arial><input type=text name=Time_Closed  value = "& objRec("Time_Closed") & "></td></tr>"	
	Response.Write"<input type=Hidden name=Time_Closed2  value ='09/22/78'>"
	End IF
	Response.Write "<tr><td><font size=1 face=Arial><b>CC:</b></td><td><input type=text name=CC_Mail></td>"
	Response.Write "<td colspan=2><font size=1 face=Arial>Send E-Mail:<input type=CheckBox name=SendeMail checked></td></tr>"
	Response.Write "<tr><td colspan=1><input type=submit value='Update Information' style='FONT-FAMILY: Arial, Geneva, Helvetica, Helv'></td><td colspan=3><font size=1 face=Arial>To CC: multiple people, use a semicolon (;) between addresses.</td></tr></table></form>"
End if

	objRec.Close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing

'**********  Show last 5 tickets *************'
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"
Set objRs=objConn.Execute("Select TOP 10 * from Calls WHERE Email =  '"& Email &"' order by Ticket_Number desc")

Response.Write "<font size=1>"
Response.Write "<table border=1 cellpadding=0 cellborder=0>"
Response.Write "<tr><td><font size=1><B>Ticket #</td>"
Response.Write "<td><font size=1><B>Call<br>Type</td>"

Response.Write "<td width=75><font size=1><B>Date Opened</td>"
Response.Write "<td><font size=1><B>Serviced<br>By</td>"
Response.Write "<td><font size=1><B>Problem</td>"

		
	While Not objRs.EOF
		If objRs("Closed") = "Yes" Then
			Response.Write "<tr bgcolor=lightgrey><td align=center><font size=1>"
		Else
			Response.Write "<tr><td align=center><font size=1>"
		End If
		Response.Write "<a href=""modifyticket.asp?Num=" & objRs("Ticket_Number") & """>" & objRs("Ticket_Number") & "</a>" & "</td>"
		If objRs("IN_PROCESS") = "Yes" Then
					Response.Write "<td bgcolor=#A8EAA8 align=center><font color=black><font size=1>"& objRs("Call_Type") & "</td>"
		End If
		If objRs("IN_PROCESS") = "MayBe" Then
					Response.Write "<td bgcolor=#FBFD04 align=center><font color=black><font size=1>"& objRs("Call_Type") & "</td>"
		End If
		If objRs("IN_PROCESS") = "No" Then
					Response.Write "<td bgcolor=#FF6666 align=center><font color=black><font size=1>"& objRs("Call_Type") & "</td>"
		End If
		Response.Write "<td><font size=1>" & objRs("Date_Opened") & "<br>"& objRs("Time_Opened") &"</td>"
		Response.Write "<td><font size=1>" & objRs("CALL_SERVICED_BY") & "</td>"						
		

		ProblemDesc = objRs("PROBLEM_DESC")
		If Len(ProblemDesc) > 100 Then
			Response.Write "<td><font size=1>" & Mid(ProblemDesc, 1, 100) & "...<a href=""modifyticket.asp?Num=" & objRs("Ticket_Number") & """>(more)</a></td>"	
		Else
			Response.Write "<td><font size=1>" & ProblemDesc & "</td>"
		End If	

		objRs.MoveNext
	Wend
objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing	
%>

</TD>
</TR>
</TABLE>
</BODY>
</HTML>

