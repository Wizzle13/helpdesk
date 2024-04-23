<!--#include file="_head.asp"-->
<%

Dim RefPage
RefPage = Request.ServerVariables("PATH_INFO") & "?" & Request.ServerVariables("QUERY_STRING")
'Response.Write RefPage
%>

<SCRIPT LANGUAGE=vbscript>
    <!--
    'SpellChecker
    ' PURPOSE: This function accepts Text data
    '     for which spell checking has to be done
    ' 
    ' Return's Spelling corrected data 
    '	
    function SpellChecker(TextValue)
    	
    	Dim objWordobject
    	Dim objDocobject 
    	Dim strReturnValue
    	'Create a new instance of word Application
    	Set objWordobject = CreateObject("word.Application")
    	objWordobject.WindowState = 2
    	objWordobject.Visible = True
    	'Create a new instance of Document
    	Set objDocobject = objWordobject.Documents.Add( , , 1, True)
    	objDocobject.Content=TextValue
    	objDocobject.CheckSpelling
    	'Return spell check completed text data
    	strReturnValue = objDocobject.Content
    	'Close Word Document 
    	objDocobject.Close False
    	'Set Document To nothing
    	Set objDocobject = Nothing
    	'Quit Word
    	objWordobject.Application.Quit True
    	'Set word object To nothing
    	Set objWordobject= Nothing
    SpellChecker=strReturnValue
    End function 
    -->
    </SCRIPT>


<script>
<!-- Script to open User selection window.
function remote(){
win2=window.open("userwindow.asp?sortby=LName","win2","width=450,height=281,titlebar=0,scrollbars=1")
win2.creator=self
}
function uploadfile(){
win3=window.open("uploads/index.asp?Num=<%=Request("Num")%>&RefPage=<%Response.Write RefPage%>","win3","width=375,height=95,titlebar=0,scrollbars=0")
win3.creator=self
}
//-->
</script>


<script language="JavaScript">
function Join_Form1_Validator(theForm)
{	
  if (theForm.FirstName.value == "")
  {
    alert("Please enter a value for the \"First Name\" field.");
    theForm.FirstName.focus();
    return (false);
  }
  
  if (theForm.LastName.value == "")
  {
    alert("Please enter a value for the \"Last Name\" field.");
    theForm.LastName.focus();
    return (false);
  }
  
   
  if (theForm.Extention.value == "")
  {
    alert("Please enter a value for the \"Extention\" field.");
    theForm.Extention.focus();
    return (false);
  }

  if (theForm.Package_Name.value == "_Unclassified_")
  {
    alert("Please enter a value for the \"Package Name\" field.");
    theForm.Package_Name.focus();
    return (false);
  }


  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.FirstName.value;
  var checkStr = theForm.LastName.value;
  var checkStr = theForm.Extention.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
   
  }
  
  return (true);
}
</script>
<%
Dim strFirstName
IF Request("num") ="" then
strFirstName=""
strLastName=""
strExtention=""
strEmail=""
strCountry=Session("Country")
Else


UserNum = Request("Num")

Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
	
objConn.Open "Helpdesk2"
sqlId ="SELECT * FROM UserInfo WHERE Id=" & UserNum & ";"

objRec.Open sqlId, objConn	

strFirstName=objRec("FirstName")
strLastName=objRec("LastName")
strExtention=objRec("Extention")
strEmail=objRec("Email")
strCountry=objRec("Country")

objRec.Close
objConn.Close
Set objRec = Nothing
Set objConn = Nothing	
End if

%>

<TD>

<form method="post" action="send.asp" name="Inputform" onsubmit='return Join_Form1_Validator(this)' name='Join_Form1'>

<table border="0" cellspacing=0 cellpadding=0>
<tr>
<td><img src=images/white_spacer.jpg></td>
<td>Date:</td><td><input type="hidden" name="Date" value=<%=FormatDateTime (Date, vbShortDate)%>><%=FormatDateTime (Date, vbShortDate)%></td>
<td>
</td>
</tr>
<tr>
<td><img src=images/white_spacer.jpg></td>
<td>Time:</td><td><input type="hidden" name="Time" value=<%=FormatDateTime (Time, vbShortTime)%>><%=FormatDateTime (Time, vbShortTime)%></td>
</tr>
<tr><td><img src=images/1_circle.jpg></td><td colspan=2><font size=2><b>Step 1: Select a user from the user list by clicking on one of the boxes below.</b></font></td></tr>
<tr>
<td><img src=images/white_spacer.jpg></td>
<td>First Name:</td><td><input type="text" name="FirstName" size="20" Value="
<%
If strFirstName <> "" then
	response.write strFirstName
else
	response.write "Click to add user"
end if
 %>" onClick="remote()"></td>
</tr>
<tr>
<td><img src=images/white_spacer.jpg></td>
<td>Last Name:</td><td><input type="text" name="LastName" size="20" Value="
<%
If strLastName <> "" then
	response.write strLastName
else
	response.write "Click to add user"
end if
%>" onClick="remote()"></td>
</tr>
<%
Response.Write"<tr><td><img src=images/white_spacer.jpg></td><td><p>Country:</td><td><select name=Country>"
	Set objConn=Server.CreateObject("ADODB.Connection")
	Set objRs = Server.CreateObject ("ADODB.Recordset")
	objConn.Open "DSN=Helpdesk2"
	Set objRs=objConn.Execute("Select * from Domains order by CountryFull")
	'sqlId ="SELECT * FROM GROUP_MEMBERS Sort by Package_Name"
	'objRs.Open sqlId, objConn	
	While Not objRs.EOF
		If strCountry=objRs("Country") then
		Response.Write"<Option value='"& objRs("Country")&"' selected>"& objRs("CountryFull")
		Else
		Response.Write"<Option value='"& objRs("Country")&"'>"& objRs("CountryFull")
		End if
		objRs.MoveNext
	Wend
	Response.Write"</select></td></tr>"
 %>	

<tr>
<td><img src=images/white_spacer.jpg></td>
<td>Contact Number:</td><td><input type="text" name="Extention" size="20" Value='<% response.write strExtention %>'></td>
</tr>
<tr>
<td><img src=images/white_spacer.jpg></td>
<td>Email:</td><td valign="bottom"><input type="Text" name="Logon_Name" size="15" Value='<% response.write strEmail %>'><Select  name=Domain>
<%Set objConn=Server.CreateObject("ADODB.Connection")
	Set objRs = Server.CreateObject ("ADODB.Recordset")
	objConn.Open "DSN=Helpdesk2"
	Set objRs=objConn.Execute("Select * from Domains order by CountryFull")
		While Not objRs.EOF
		If strCountry=objRs("Country") then
			Response.Write "<Option selected>"& objRs("Domain") &""
		Else
			Response.Write "<Option>"& objRs("Domain") &""
		End iF	
			objRs.MoveNext
	Wend
	objConn.Close
	
	Set objConn = Nothing
		
Response.Write "</select></td></tr>"

Response.Write "<tr><td><img src=images/2_circle.jpg></td><td colspan=2><font size=2><b>Step 2: If you want to attach a file to this ticket, do it next. If you have multiple files to upload, please zip them together and upload the zipped file.</b></font></td></tr>"

If Request("FileName") <> "" Then
	Response.Write "<tr><td><img src=images/white_spacer.jpg></td><td>File<br>Attachment:</td><td><a href='/uploads/data/" & Request("FileName") & "' target='new'>" & Request("FileName") & "</a>&nbsp;"
	Response.Write "<input type=hidden name='Upload_File_Location' value='/uploads/data/" & Request("FileName") & "'>"
Else
	Response.Write "<tr><td><img src=images/white_spacer.jpg></td><td>File<br>Attachment:</td><td>&nbsp;"
End If
Response.Write "<input type=button onClick=uploadfile() value='Attach File'></td></tr>"

Response.Write "<tr><td><img src=images/3_circle.jpg></td><td colspan=2><font size=2><b>Step 3: Complete the rest of the information and submit.</b></font></td></tr>"

Response.Write "<tr><td><img src=images/white_spacer.jpg></td><td>Call Type: </td><td><select name='Call_Type'><option selected>HELP<option>REQUEST<option>HELP-3rd Party</select></td></tr>"


	Dim sqlId
	Dim objConn
	Dim objRs
	Response.Write"<tr><td><img src=images/white_spacer.jpg></td><td><p>Package Name:</td><td><select name=Package_Name>"
	Set objConn=Server.CreateObject("ADODB.Connection")
	Set objRs = Server.CreateObject ("ADODB.Recordset")
	objConn.Open "DSN=HelpDesk2"
	Set objRs=objConn.Execute("Select * from GROUP_MEMBERS order by Package_name")
	'sqlId ="SELECT * FROM GROUP_MEMBERS Sort by Package_Name"
	'objRs.Open sqlId, objConn	
	While Not objRs.EOF
	Response.Write"<Option value='"& objRs("Package_name")&"'>"& objRs("Package_Name")
		objRs.MoveNext
	Wend
	Response.Write"</select></td></tr>"
	
	objConn.Close
	
	Set objConn = Nothing	
	
	
	Response.Write"<tr><td><img src=images/white_spacer.jpg></td><td>Person Working Call:</td><td><select name=Call_Serviced_By>"
	Set objConn=Server.CreateObject("ADODB.Connection")
	objConn.Open "DSN=HelpDesk2"
	Set objRs=objConn.Execute("Select * from Workers where Active_Worker='Yes' order by First_Name, Last_Name")
	sqlId ="SELECT * FROM Workers;"
	While Not objRs.EOF
	Response.Write"<Option value="& objRs("Worker_ID")&">"& objRs("First_Name")&" "& objRs("Last_Name")
		objRs.MoveNext
	Wend

	Response.Write"</td>"
	Response.Write"</tr>"
Response.Write "<tr><td><img src=images/white_spacer.jpg></td><td>Priority:</td><td><input type=radio value =1 name=Priority>High <input type=radio value=2 name=Priority Checked>Normal <input type=radio value=3 name=Priority>Low</td></tr>"

	Response.Write "<tr><td><img src=images/white_spacer.jpg></td><td>SAP Priority: </td><td><select name=tickpriority>" & vbcrlf
		i = 1
		While i <> 21
			Response.Write "<option value=" & i & ">" & i & vbcrlf
			i = i + 1
		Wend
		Response.Write "<option value=99 selected>99"
		Response.Write "</select></td></tr>"  & vbcrlf
		strDate=FormatDateTime (Date, vbShortDate)
		
		Response.Write "<input type=hidden name='SAP_Transport_Date' value='12/31/2049'>"
		
'		Response.Write "<tr><td><img src=images/white_spacer.jpg></td><td>Expected Production Transport Date:</td>"
'		Response.Write "<td><select name=SAP_Transport_Date>"
'		Response.Write "<option value=N/A selected>N/A"
'		Set objRs1=objConn.Execute("Select * from SAP_Transports")
'		sqlId ="SELECT * FROM SAP_Transports;"
'	While Not objRs1.EOF
'		If objRs1("Transport_Date") > Date then
'			Response.Write"<Option value="& objRs1("Transport_Date")&">"& objRs1("Transport_Date")
'		End if
'		objRs1.MoveNext
'	Wend
'		Response.Write "</td></tr>"
	Response.Write "<tr><td><img src=images/white_spacer.jpg></td><td>SAP Module: </td><td><select name=SAP_Module>" & vbcrlf	
			Response.Write "<option value=N/A selected>N/A"
			Response.Write "<option value=SD>SD"
			Response.Write "<option value=MM>MM"
			Response.Write "<option value=PP>PP"
			Response.Write "<option value=FI/CO>FI/CO"
			Response.Write "<option value=Basis>Basis"
			Response.Write"<Option Value=WM>WM"
			Response.Write"<Option Value=BW>BW"
			Response.Write"<Option Value=APO>APO"

	Response.Write "<tr><td><img src=images/white_spacer.jpg></td><td colspan=1><p>Projected Complete Date:</td>"
	Response.Write "<td><font size=1 face=Arial><input type=Text name=Projected_Complete_Date></td></tr>"
	Response.Write "<input type=Hidden name=Projected_Complete_Date2 value='12/31/2049'>"		
		
	Response.Write "<tr><td><img src=images/white_spacer.jpg></td><td colspan=2>"
	Response.Write"Description:<BR>"
	Response.Write"<textarea cols=50 rows=15 name=problem></textarea><br></td></tr>"
	Response.Write"</table>"
	'Response.Write "<input type=button onClick='Inputform.problem.value=SpellChecker(Inputform.problem.value)' value='SC'>"

 %>
<p>

<img src=images/white_spacer.jpg>
<input type="submit" value="Save" name="B1"> 
<input type="submit" value="Save & New" name="B1">
<input type="submit" value="Save & View" name="B1">
</form>
</TD>
</TR>
</TABLE>
</TD>
</TR>
</TABLE>

</BODY>
</HTML>


