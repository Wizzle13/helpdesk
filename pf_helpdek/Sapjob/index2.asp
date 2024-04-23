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
<style>
input {FONT-FAMILY: Arial, Geneva, Helvetica, Helv; FONT-SIZE: xx-small; FONT-STYLE: normal}
select {FONT-FAMILY: Arial, Geneva, Helvetica, Helv; FONT-SIZE: xx-small; FONT-STYLE: normal}
</style>


<script language="JavaScript">
function FrontPage_Form2a_Validator(theForm)
{	
  if (theForm.ProgramName1.value == "")
  {
    alert("Please enter a value for the \"Program Name\" field.");
    theForm.ProgramName1.focus();
    return (false);
  }
  
  if (theForm.VariantName1.value == "")
  {
    alert("Please enter a value for the \"Variant Name\" field.");
    theForm.VariantName1.focus();
    return (false);
  }
  
  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.ProgramName1.value;
  var checkStr = theForm.VariantName1.value;
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

function FrontPage_Form2b_Validator(theForm)
{	
  if (theForm.CommandName.value == "")
  {
    alert("Please enter a value for the \"Command Name\" field.");
    theForm.CommandName.focus();
    return (false);
  }
  
  if (theForm.Parameters.value == "")
  {
    alert("Please enter a value for the \"Parameters Name\" field.");
    theForm.Parameters.focus();
    return (false);
  }
  
  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.CommandName.value;
  var checkStr = theForm.Parameters.value;
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

function FrontPage_Form2c_Validator(theForm)
{	
  if (theForm.ProgramName.value == "")
  {
    alert("Please enter a value for the \"Program Name\" field.");
    theForm.ProgramName.focus();
    return (false);
  }
  
  if (theForm.Parameter.value == "")
  {
    alert("Please enter a value for the \"Parameter Name\" field.");
    theForm.Parameter.focus();
    return (false);
  }
  
  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.ProgramName.value;
  var checkStr = theForm.Parameter.value;
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

<P>&nbsp;</P>
<font face="Arial" size="4">SAP Background Job Form</font>
<%
	' Declare all information from the form as Session variables.
	Session("Name") = Request.Form("Name")
	Session("JobName") = Request.Form("JobName")
	Session("JobClass") = Request.Form("JobClass")
	Session("TargetHost") = Request.Form("TargetHost")
	Session("FunctionalArea") = Request.Form("FunctionalArea")
	Session("StepType") = Request.Form("StepType")

	If Session("StepType") = "ABAP Program" Then
	Response.Write "<form method='post' action='index3.asp' onsubmit='return FrontPage_Form2a_Validator(this)' name='FrontPage_Form2a'><table>"
	End If
	If Session("StepType") = "External Command" Then	
	Response.Write "<form method='post' action='index3.asp' onsubmit='return FrontPage_Form2b_Validator(this)' name='FrontPage_Form2b'><table>"
	End If
	If Session("StepType") = "External Program" Then
	Response.Write "<form method='post' action='index3.asp' onsubmit='return FrontPage_Form2c_Validator(this)' name='FrontPage_Form2c'><table>"
	End If

	Response.Write "<tr><td><font size='2'>Submitted By:</td><TD><font size='2'>" & Session("Name") & "</td></tr>"
	Response.Write "<tr><td><font size='2'>Job Name:</td><TD><font size='2'>" & Session("JobName") & "</td></tr>"
	Response.Write "<tr><td><font size='2'>Job Class:</td><TD><font size='2'>" & Session("JobClass") & "</td></tr>"
	Response.Write "<tr><td><font size='2'>Target Host:</td><TD><font size='2'>" & Session("TargetHost")& "</td></tr>"
	Response.Write "<tr><td><font size='2'>Functional Area:</td><TD><font size='2'>" & Session("FunctionalArea")& "</td></tr>"
	Response.Write "<tr><td><font size='2'>Step Type:</td><TD><font size='2'>" & Session("StepType")& "</td></tr>"
	
If Session("StepType") = "ABAP Program" Then
	%>
	<!--#include file="ABAPProgram.asp"-->
	<%
End If

If Session("StepType") = "External Command" Then
	%>
	<!--#include file="ExternalCommand.asp"-->
	<%
End If

If Session("StepType") = "External Program" Then
	%>
	<!--#include file="ExternalProgram.asp"-->
	<%
End If
%>
<tr>
		<td valign="top"><font size="2">Start Time:</td>
		<td><!--<font size=1><input type="radio" name="StartTime" value="Immediate">Immediate<br>-->
			<input type="radio" name="StartTime" value="Date/Time" checked>Date/Time<br>
			<input type="radio" name="StartTime" value="After Job">After Job<BR>
			<input type="radio" name="StartTime" value="After Event">After Event<br>
			<input type="radio" name="StartTime" value="At Operation Mode">At Operation Mode<br></font>
		</td>
	</tr>
</table>
<input type="button" onclick="javascript:history.go(-1)" value="<<Back">
<input type="submit" value="Continue >>">

</form>

</TD>
</TR>
</TABLE>
</BODY>
</HTML>


