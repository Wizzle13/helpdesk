<!--#include file="../_validsession.asp"-->
<!--
Programer: Chris Burton    
Date Started: 7-9-01

Description:
This page checks the  database for users First and Last name
then uses it to display a hello message to the user.

-->
<!--#include file="../_head.asp"-->
<%
Response.Write "</td>"
Response.Write "<td valign=Top>"
%>

<script language="JavaScript">
function FrontPage_Form1_Validator(theForm)
{
	
 if (theForm.Name.value == "")
  {
    alert("Please enter a value for the \"Submitted By\" field.");
    theForm.Name.focus();
    return (false);
  }
 
  if (theForm.JobName.value == "")
  {
    alert("Please enter a value for the \"Job Name\" field.");
    theForm.JobName.focus();
    return (false);
  }
  
  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.JobName.value;
  var allValid = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("There cannot be any spaces in the \"Job Name\" field.  Please replace any spaces with an underscore( _ ) .");
    theForm.JobName.focus();
    return (false);
  }
  
  return (true);
}
</script>

<font face='Arial' size='4'>SAP Background Job Form</font>
<form method='post' action='index2.asp' onsubmit='return FrontPage_Form1_Validator(this)' name='FrontPage_Form1'>

<table><tr>
	<td><font size='2'>Submitted By:</font></td>
		<td><input name='Name' type='Hidden' value=<%=session("Email") %>><%=session("Email") %></td>
		
	</tr>

	<tr>
		<td><font size="2">Job Name:</font></td>
		<td><input name="JobName" type="text" size="40"></td>
	</tr>
	<tr>
		<td><font size="2">Job Class:</td>
		<td><select name="JobClass">
				<option value="B">B
				<option value="C" selected>C
		</select></td>
	</tr>
	<tr>
		<TD><font size="2">Target Host:</td>
		<td><select name="TargetHost">
				<option selected>APP4PRD
				<option>APP5PRD
				<option>CLUENT
		</select></td>
	</tr>
	<tr>
		<td><font size="2">Functional<br>Area:</td>
		<td><select name="FunctionalArea">
				<option value="CO">CO - Controlling
				<option value="FI">FI - Finance
				<option value="IT">IT - Basis/Programming
				<option value="MM">MM - Materials Management
				<option value="PP">PP - Production Planning & Execution
				<option value="SD">SD - Sales & Distribution
		</select>
		</td>
	</tr>
	<tr>
		<td valign="top"><font size="2">Step Type:</td>
		<td><font size="2"><input type="radio" name="StepType" value="ABAP Program" checked>ABAP Program<br>
			<input type="radio" name="StepType" value="External Command">External Command<br>
			<input type="radio" name="StepType" value="External Program">External Program</font>
		</td>
	</tr>
</table>
<input type="submit" value="Continue >>">

</form>

</TD>
</TR>
</TABLE>
</BODY>
</HTML>

