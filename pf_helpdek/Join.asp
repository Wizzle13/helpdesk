<!--
Programer: Chris Burton    
Date Started: 7-9-01

Description:
This Page is the Join page for the User interface for the Helpdesk.
This page uses JavaScript to check that all the information is entered.

-->
<HTML>
<HEAD>
  <TITLE>Information Technology's Intranet Page</TITLE>

<script language="JavaScript">
function Join_Form1_Validator(theForm)
{	
  if (theForm.First_Name.value == "")
  {
    alert("Please enter a value for the \"First Name\" field.");
    theForm.First_Name.focus();
    return (false);
  }
  
  if (theForm.Last_Name.value == "")
  {
    alert("Please enter a value for the \"Last Name\" field.");
    theForm.Last_Name.focus();
    return (false);
  }
  if (theForm.Extention.value == "")
  {
    alert("Please enter a value for the \"Extention\" field.");
    theForm.Extention.focus();
    return (false);
  }
    if (theForm.Logon_Name.value == "")
  {
    alert("Please enter a value for the \"Email\" field.");
    theForm.Logon_Name.focus();
    return (false);
  }

  if (theForm.Department.value == "None")
  {
    alert("You Must Select Your Department.");
    theForm.Department.focus();
    return (false);
  }
  
    if (theForm.Country.value == "None")
  {
    alert("You Must Select Your Country.");
    theForm.Country.focus();
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
  var checkStr = theForm.First_Name.value;
  var checkStr = theForm.Last_Name.value;
  var checkStr = theForm.Extention.value;
  var checkStr = theForm.Logon_Name.value;
  var checkStr = theForm.Password.value;
  var checkStr = theForm.Password2.value;
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
</HEAD>
<body>



<font size=5>Pure Fishing Online Help Desk Join Screen</font>


<P>


<form method="post" action="Joinsend.asp" name="Inputform" onsubmit='return Join_Form1_Validator(this)' name='Join_Form1'>

<table border="1">
<tr>
<td>First Name:</td><td><input type="text" name="First_Name" size="20"></td>
</tr>
<tr>
<td>Last Name:</td><td><input type="text" name="Last_Name" size="20"></td>
</tr>
<tr>
<td>Contact Number:<font color="Red">*</font></td><td><input type="text" name="Extention" size="20"></td>
</tr>
<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
Set objRs=objConn.Execute("Select * from Departments Order by Department_Name")
 %>

<tr>
<td>Department:</td><td><Select  name=Department> 
<option value=None selected>Select Your Department
<option value=None>-----------------
<%
		While Not objRs.EOF
			Response.Write "<Option>"& objRs("Department_Name") &""
			objRs.MoveNext
	Wend
objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing
	
 %>		
 </select></td>
</tr>
<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
Set objRs=objConn.Execute("Select * from Domains")
 %>
<tr>
<td>Email:</td><td><input type="Text" name="Logon_Name" size="15"><Select  name=Domain>
<%
		While Not objRs.EOF
			Response.Write "<Option>"& objRs("Domain") &""
			objRs.MoveNext
	Wend
objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing
	
 %>		
 </select></td>
</tr>
<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
Set objRs=objConn.Execute("Select * from Domains Order by CountryFull")
 %>
<tr>
<td>Country:</td><td><Select  name=Country><Option Selected Value=None>Select Your Country<Option Value=None>---------------------
<%
		While Not objRs.EOF
			Response.Write "<Option Value="& objRs("Country") &">"& objRs("CountryFull") &""
			objRs.MoveNext
	Wend
objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing
	
 %>		
 </select></td>
</tr>
<tr>
<td>Password:<font color="Red">**</font></td><td><input type="Password" name="Password" size="20"></td>
</tr>
<tr>
<td> Confirm Password:<font color="Red">**</font></td><td><input type="Password" name="Password2" size="20"></td>
</tr>

</table>
<input type="Hidden" name="UserAdded" Value="0">
<br>
<font color="Red">*</font>Enter the main number where you can be reached.<BR>
<font color="Red">**</font>This is different and separate from your Windows and SAP passwords.
<p>
<input type="submit" value="Send" name="B1"> 
<input type="reset"value="Reset" name="B2">
</form>
</TD>
</TR>
</TABLE>
</BODY>
</HTML>

