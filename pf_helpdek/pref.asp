<!--#include file="_head.asp"-->
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
  if (theForm.Department.value == "")
  {
    alert("Please enter a value for the \"Department\" field.");
    theForm.Department.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.FirstName.value;
  var checkStr = theForm.LastName.value;
  var checkStr = theForm.Extention.value;
  var checkStr = theForm.Password.value;
  var checkStr = theForm.WorkerID.value;
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

<font size=1 face="Arial">
<%
	Dim sqlId
	Dim objConn
	Dim objRec
	Dim rsId
	Dim rsName

	
	Response.Write "<td valign=top>"
	Response.Write ""
	Response.Write "<form method=post action=save2.asp?Num=" & Session("UID") & "  onsubmit='return Join_Form1_Validator(this)' name='Join_Form1'>"
	Response.Write "<table border=0><tr>"
	Response.Write "<td>First Name: </td>"
	Response.Write "<td><input type=text name=FirstName value=" & Session("FirstName")&"></td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Last Name: </td>"
	Response.Write "<td><input type=text name=LastName value=" & Session("LastName")&"></td>"
	Response.Write "</tr><tr>"

	Response.Write "<td>Email:</td>"
	Response.Write "<td><input type=text name=Email size=10 value=" & Session("Email") &">"
	Set objConn=Server.CreateObject("ADODB.Connection")
	objConn.Open "DSN=Helpdesk2"
	Set objRs=objConn.Execute("Select * from Domains")
	Response.Write "<Select name=Domain>"
	While Not objRs.EOF
		If Session("Country")=objRs("Country") then
			Response.Write "<Option Selected Value'"&objRs("Domain")&"'>"& objRs("Domain") &""
		else
			Response.Write "<Option Value='"&objRs("Domain")&"'>"& objRs("Domain") &""
		End if	
			objRs.MoveNext
	Wend
	objRs.close
	Set objRs= Nothing
	objConn.Close
	set objConn=nothing
	Response.Write "</td>"
	Response.Write "</tr><tr>"
Response.Write "<td>Country:</td>"
	Response.Write "<td>"
	Set objConn=Server.CreateObject("ADODB.Connection")
	objConn.Open "DSN=Helpdesk2"
	Set objRs=objConn.Execute("Select * from Domains")
	Response.Write "<Select name=Country>"
	While Not objRs.EOF
		If Session("Country")=objRs("Country") then
			Response.Write "<Option Selected Value='"&objRs("Country")&"'>"& objRs("CountryFull") &""
		else
			Response.Write "<Option Value='"&objRs("Country")&"'>"& objRs("CountryFull") &""
		End if	
			objRs.MoveNext
	Wend
	objRs.close
	Set objRs= Nothing
	objConn.Close
	set objConn=nothing
	Response.Write "</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Contact Number:</td>"
	Response.Write "<td><input type=text name=Extention value='" & Session("Extention") &"'></td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Department:</td>"
	Response.Write "<td><select Name=Department><Option value='" & Session("Department") &"'>" & Session("Department")
	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")

	
	objConn.Open "Helpdesk2"


	Set objRs=objConn.Execute("Select * from Departments order by Department_name")
	sqlId ="SELECT * FROM Workers;"
	While Not objRs.EOF
	Response.Write"<Option value='"& objRs("Department_Name")&"'>"& objRs("Department_Name")
		objRs.MoveNext
	Wend
	
	Response.Write  "</td></tr><tr>"
	objRs.Close
	objConn.Close
	Set objRs = Nothing
	Set objConn = Nothing
	Response.Write "<td><input type=Hidden name=Password value=" & Session("Password") &">"
	Response.Write "<input type=Hidden name=Status value=" & Session("Status") &">"
	Response.Write "<input type=Hidden name=WorkerID value=" & Session("WorkID") &"></td>"
	Response.Write "</tr>"
	

	
	Response.Write "<tr><td colspan=2><input type=Submit  value='Update Information'></td></tr></table>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

		
%>
<p>



</TD>
</TR>
</TABLE>
</BODY>
</HTML>

