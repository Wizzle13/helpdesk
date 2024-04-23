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

  if (theForm.Password.value == "")
  {
    alert("Please enter a value for the \"Password\" field.");
    theForm.Password.focus();
    return (false);
  }
   if (theForm.Password2.value == "")
  {
    alert("Please enter a value for the \"Password2\" field.");
    theForm.Password2.focus();
    return (false);
  }
   if (theForm.Password.value != theForm.Password2.value)
  {
    alert("Passwords do not match. Please retype passwords.");
    theForm.Password.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.FirstName.value;
  var checkStr = theForm.LastName.value;
  var checkStr = theForm.Extention.value;
  var checkStr = theForm.Password.value;
  var checkStr = theForm.Password2.value;
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

	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")

	
	objConn.Open "Helpdesk2"

	
	sqlId ="SELECT * FROM UserInfo WHERE ID=" & Request("Num") & ";"

	objRec.Open sqlId, objConn	

	Response.Write "<td>"
	Response.Write ""
	Response.Write "<P><a href=DeleteMember.asp?Num=" & objRec("ID") & ">Delete</a>"
	Response.Write "<form method=post action=save.asp?Num=" & Request("Num") & " onsubmit='return Join_Form1_Validator(this)' name='Join_Form1'>"
	Response.Write "<table border=0><tr>"
	Response.Write "<td>First Name: </td>"
	Response.Write "<td><input type=text name=FirstName value=" & objRec("FirstName")&"></td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Last Name: </td>"
	Response.Write "<td><input type=text name=LastName value=" & objRec("LastName")&"></td>"
	Response.Write "</tr><tr>"

	Response.Write "<td>Email:</td>"
	Response.Write "<td><input type=text name=Email value=" & objRec("Email") &" size=10>"
	Set objRs=objConn.Execute("Select * from Domains")
	Response.Write "<Select name=Domain>"
	While Not objRs.EOF
		If ObjRec("Country")=objRs("Country") then
			Response.Write "<Option Selected Value='"&objRs("Domain")&"'>"& objRs("Domain") &""
		else
			Response.Write "<Option Value='"&objRs("Domain")&"'>"& objRs("Domain") &""
		End if	
			objRs.MoveNext
	Wend
	objRs.close
	Set objRs= Nothing
	
	Response.Write "</td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Country:</td>"
	Response.Write "<td>"
	Set objRs=objConn.Execute("Select * from Domains")
	Response.Write "<Select name=Country>"
	While Not objRs.EOF
		If ObjRec("Country")=objRs("Country") then
			Response.Write "<Option Selected Value='"&objRs("Country")&"'>"& objRs("CountryFull") &""
		else
			Response.Write "<Option Value='"&objRs("Country")&"'>"& objRs("CountryFull") &""
		End if	
			objRs.MoveNext
	Wend
	objRs.close
	Set objRs= Nothing
	Response.Write "</td>"
	Response.Write "</tr><tr>"

	
	Response.Write "<td>Extention:</td>"
	Response.Write "<td><input type=text name=Extention value='" & objRec("Extention") &"'></td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Department:</td>"
	Response.Write "<td><select Name=Department><Option value='" & objRec("Department") &"'>" & objRec("Department")
	

	Set objRs=objConn.Execute("Select * from Departments order by Department_name")
	sqlId ="SELECT * FROM Workers;"
	While Not objRs.EOF
	Response.Write"<Option value='"& objRs("Department_Name")&"'>"& objRs("Department_Name")
		objRs.MoveNext
	Wend
	
	Response.Write  "</td></tr><tr>"
	Response.Write "<td>Password:</td>"
	Response.Write "<td><input type=Password name=Password value=" & objRec("Password") &"></td>"
	Response.Write "</tr><tr>"
	Response.Write "<td>Confirm Password:</td>"
	Response.Write "<td><input type=Password name=Password2 value=" & objRec("Password") &"></td>"
	If objRec("Status") = "Administrator" Then
		Response.Write "<tr><td>Status:</td><td><select name=Status><option selected>Administrator<option>Manager<option>Worker<option>Member</select><BR></td></tr>"
	End If	
	If objRec("Status") = "Manager" Then
		Response.Write "<tr><td>Status:</td><td><select name=Status><option>Administrator<option selected>Manager<option>Worker<option>Member</select><BR></td></tr>"
	End If
	If objRec("Status") = "Worker" Then
		Response.Write "<tr><td>Status:</td><td><select name=Status><option>Administrator<option>Manager<option selected>Worker<option>Member</select><BR></td></tr>"
	End If

	If objRec("Status") = "Member" Then
		Response.Write "<tr><td>Status:</td><td><select name=Status><option>Administrator<option>Manager<option>Worker<option selected>Member</select><BR></td></tr>"
	End If

	Response.Write "<td>Worker ID:</td>"
	Response.Write "<td><input type=text name=WorkerID value=" & objRec("WorkerID") &"></td>"
	Response.Write "</tr>"
	

	
	Response.Write "<tr><td colspan=2><input type=Submit  value=Update Information  style='FONT-FAMILY: Arial, Geneva, Helvetica, Helv'></td></tr></table>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

	objRec.Close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing	
	
%>
<p>



</TD>
</TR>
</TABLE>
</BODY>
</HTML>

