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
function FrontPage_Form3a_Validator(theForm)
{
  if (theForm.RecipientName.value == "")
  {
    alert("Please enter a value for the \"Recipient Name\" field.");
    theForm.RecipientName.focus();
    return (false);
  }
  
  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.RecipientName.value;

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
    alert("There cannot be any spaces in the \"Recipient Name\" field.  Please replace any spaces with an _.");
    theForm.RecipientName.focus();
    return (false);
  }
  
   return (true);
}

function FrontPage_Form3b_Validator(theForm)
{
  if (theForm.StartHour.value == "")
  {
    alert("Please enter a value for the \"Needed Time\" field.");
    theForm.StartHour.focus();
    return (false);
  }
  if (theForm.StartMonth.value == "")
  {
    alert("Please enter a value for the \"Needed Month\" field.");
    theForm.StartMonth.focus();
    return (false);
  }

 if (theForm.StartDate.value == "")
  {
    alert("Please enter a value for the \"Needed Date\" field.");
    theForm.StartDate.focus();
    return (false);
  }
  
 
   return (true);
}
function FrontPage_Form3c_Validator(theForm)
{
  if (theForm.AfterJob.value == "")
  {
    alert("Please enter a value for the \"  After Job\" field.");
    theForm.  AfterJob.focus();
    return (false);
  }
  if (theForm.RecipientName.value == "")
  {
    alert("Please enter a value for the \"Recipient Name\" field.");
    theForm.RecipientName.focus();
    return (false);
  }
  
  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.AfterJob.value;
  var checkStr = theForm.RecipientName.value;
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
    alert("There cannot be any spaces in the \"Recipient Name\" field.  Please replace any spaces with an _.");
    theForm.RecipientName.focus();
    return (false);
  }
  
   return (true);
}
function FrontPage_Form3c_Validator(theForm)
{
  if (theForm.AfterJob.value == "")
  {
    alert("Please enter a value for the \"  After Job\" field.");
    theForm.  AfterJob.focus();
    return (false);
  }
 if (theForm.RecipientName.value == "")
  {
    alert("Please enter a value for the \"Recipient Name\" field.");
    theForm.RecipientName.focus();
    return (false);
  } 
  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.AfterJob.value;
  var checkStr = theForm.RecipientName.value;
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
    alert("There cannot be any spaces in the \"Recipient Name\" field.  Please replace any spaces with an _.");
    theForm.RecipientName.focus();
    return (false);
  }
   return (true);
}

function FrontPage_Form3d_Validator(theForm)
{
  if (theForm.EventName.value == "")
  {
    alert("Please enter a value for the \"EventName\" field.");
    theForm.EventName.focus();
    return (false);
  }
  
 if (theForm.Parameter.value == "")
  {
    alert("Please enter a value for the \"Parameter\" field.");
    theForm.Parameter.focus();
    return (false);
  }
   
  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.EventName.value;
  var checkStr = theForm.Parameter.value;
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
 
  
   return (true);
}
function FrontPage_Form3e_Validator(theForm)
{
  if (theForm.ModeName.value == "")
  {
    alert("Please enter a value for the \"Mode Name\" field.");
    theForm.ModeName.focus();
    return (false);
  }
  
 if (theForm.RecipientName.value == "")
  {
    alert("Please enter a value for the \"Recipient Name\" field.");
    theForm.RecipientName.focus();
    return (false);
  }
   
  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.ModeName.value;
  var checkStr = theForm.RecipientName.value;
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
 
  
   return (true);
}

  </script>

<P>&nbsp;</P>
<font face="Arial" size="4">SAP Background Job Form</font>
<%
'***********************************************
'* Turn all variables into Session variables.  *
'***********************************************
Session("StartTime") = Request.Form("StartTime")

'*****************************************************************
'*Begin sorting out Start Time values, and display code for each.*
'*The following case statement will tell the page what to display*
'*in the instance of each variable.								 *
'*****************************************************************
	Select Case Session("StartTime")
		'******************************************
		'* This is the code to execute if         *
		'* StartTime is "Immediate"               *
		'******************************************	
		Case "Immediate"
			Response.Write "<form method='post' action='display.asp' onsubmit='return FrontPage_Form3a_Validator(this)' name='FrontPage_Form3a'><table>"

			Response.Write "<tr><td><font size='2'>Job Name:</td><TD><font size='2'>" & Session("JobName")& "</td></tr>"
			Response.Write "<tr><td><font size='2'>Job Class:</td><TD><font size='2'>" & Session("JobClass")& "</td></tr>"
			Response.Write "<tr><td><font size='2'>Target Host:</td><TD><font size='2'>" & Session("TargetHost")& "</td></tr>"
			Response.Write "<tr><td><font size='2'>Functional Area:</td><TD><font size='2'>" & Session("FunctionalArea")& "</td></tr>"
			Response.Write "<tr><td><font size='2'>Step Type:</td><TD><font size='2'>" & Session("StepType")& "</td></tr>"
	
			Select Case Session("StepType")
				Case "ABAP Program"
					Session("ProgramName1") = Request.Form("ProgramName1")
					Session("VariantName1") = Request.Form("VariantName1")
					
					Response.Write "<tr><td><font size='2'>Program Name 1:</td><TD><font size='2'>" & Session("ProgramName1")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 1:</td><TD><font size='2'>" & Session("VariantName1")& "</td></tr>"
					Session("Language") = Request.Form("Language")

					Response.Write "<tr><td><font size='2'>Language:</td><TD><font size='2'>" & Session("Language")& "</td></tr>"
				Case "External Command"
					Session("CommandName") = Request.Form("CommandName")
					Session("Parameters") = Request.Form("Parameters")
					Session("OperatingSystem") = Request.Form("OperatingSystem")
					
					Response.Write "<tr><td><font size='2'>Command Name:</td><TD><font size='2'>" & Session("CommandName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameters:</td><TD><font size='2'>" & Session("Parameters")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Operating System:</td><TD><font size='2'>" & Session("OperatingSystem")& "</td></tr>"
				Case "External Program"
					Session("ProgramName") = Request.Form("ProgramName")
					Session("Parameter") = Request.Form("Parameter")
					
					Response.Write "<tr><td><font size='2'>ProgramName 1:</td><TD><font size='2'>" & Session("ProgramName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameter 1:</td><TD><font size='2'>" & Session("Parameter")& "</td></tr>"
			End Select

			Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "</td></tr>"
	
			Response.Write "<tr><td><font size='2'>Recipient User Name:</td><td><font size='2'><input type=Text name=RecipientName value=></td></tr>"
			Response.Write "<tr><td valign=top><font size='2'>Type of Output Desired:</td><td><font size='2'>"
			' Check boxes for Output Type
			Response.Write "<input type=checkbox name=OutputType value='Copy Required' checked>Copy Required<BR>"
			Response.Write "<input type=checkbox name=OutputType value='Copy'>Copy<BR>"
			Response.Write "<input type=checkbox name=OutputType value='No Printing'>No Printing<BR>"
			Response.Write "<input type=checkbox name=OutputType value='Express'>Express<BR>"
			Response.Write "<input type=checkbox name=OutputType value='To Do'>To Do<BR>"
			Response.Write "<input type=checkbox name=OutputType value='Blind Copy'>Blind Copy<BR>"
			Response.Write "<input type=checkbox name=OutputType value='No Forwarding'>No Forwarding<BR>"
			Response.Write "<input type=checkbox name=OutputType value='Repeat Send'>Repeat Send<BR>"
			'*************************************************
		
		
		Case "Date/Time"
			Response.Write "<form method='post' action='index4.asp' onsubmit='return FrontPage_Form3b_Validator(this)' name='FrontPage_Form3b'><table>"
			Response.Write "<tr><td><font size='2'>Subbmitted By:</td><TD><font size='2'>" & Session("Name")& "<input type =hidden value='" & Session("Name")& "' name=Name></td></tr>"
			Response.Write "<tr><td><font size='2'>Job Name:</td><TD><font size='2'>" & Session("JobName")& "<input type =hidden value='" & Session("JobName")& "' name=JobName></td></tr>"
			Response.Write "<tr><td><font size='2'>Job Class:</td><TD><font size='2'>" & Session("JobClass")& "<input type =hidden value='" & Session("JobClass")& "' name=JobClass></td></tr>"
			Response.Write "<tr><td><font size='2'>Target Host:</td><TD><font size='2'>" & Session("TargetHost")& "<input type =hidden value='" & Session("TargetHost")& "' name=TargetHost></td></tr>"
			Response.Write "<tr><td><font size='2'>Functional Area:</td><TD><font size='2'>" & Session("FunctionalArea")& "<input type =hidden value='" & Session("FunctionalArea")& "' name=FunctionalArea></td></tr>"
			Response.Write "<tr><td><font size='2'>Step Type:</td><TD><font size='2'>" & Session("StepType")& "<input type =hidden value='" & Session("StepType")& "' name=StepType></td></tr>"
	
			Select Case Session("StepType")
				Case "ABAP Program"
					Session("ProgramName1") = Request.Form("ProgramName1")
					Session("VariantName1") = Request.Form("VariantName1")
					Session("ProgramName2") = Request.Form("ProgramName2")
					Session("VariantName2") = Request.Form("VariantName2")
					Session("ProgramName3") = Request.Form("ProgramName3")
					Session("VariantName3") = Request.Form("VariantName3")
					Session("ProgramName4") = Request.Form("ProgramName4")
					Session("VariantName4") = Request.Form("VariantName4")
					Session("ProgramName5") = Request.Form("ProgramName5")
					Session("VariantName5") = Request.Form("VariantName5")

					Session("Language") = Request.Form("Language")
					
					Response.Write "<tr><td><font size='2'>Program Name 1:</td><TD><font size='2'>" & Session("ProgramName1")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 1:</td><TD><font size='2'>" & Session("VariantName1")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 2:</td><TD><font size='2'>" & Session("ProgramName2")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 2:</td><TD><font size='2'>" & Session("VariantName2")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 3:</td><TD><font size='2'>" & Session("ProgramName3")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 3:</td><TD><font size='2'>" & Session("VariantName3")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 4:</td><TD><font size='2'>" & Session("ProgramName4")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 4:</td><TD><font size='2'>" & Session("VariantName4")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 5:</td><TD><font size='2'>" & Session("ProgramName5")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 5:</td><TD><font size='2'>" & Session("VariantName5")& "</td></tr>"


					Response.Write "<tr><td><font size='2'>Language:</td><TD><font size='2'>" & Session("Language")& "</td></tr>"
				Case "External Command"
					Session("ProgramName") = Request.Form("ProgramName")
					Session("Parameters") = Request.Form("Parameters")
					Session("OperatingSystem") = Request.Form("OperatingSystem")
					
					Response.Write "<tr><td><font size='2'>Command Name:</td><TD><font size='2'>" & Session("CommandName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameters:</td><TD><font size='2'>" & Session("Parameters")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Operating System:</td><TD><font size='2'>" & Session("OperatingSystem")& "</td></tr>"
				Case "External Program"
					Session("ProgramName") = Request.Form("ProgramName")
					Session("Parameter") = Request.Form("Parameter")
					
					Response.Write "<tr><td><font size='2'>ProgramName:</td><TD><font size='2'>" & Session("ProgramName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameter:</td><TD><font size='2'>" & Session("Parameter")& "</td></tr>"
			End Select
			
			Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "<input type=hidden value='" & Session("StartTime")& "' name=StartTime></td></tr>"
			
			%>
			<!--#include file="StartDateTime.asp"-->
			<%
	
			Response.Write "<tr><td><font size='2'>Run Time:</td><td><font size='2'><input type=checkbox name=RunTime value='PeriodicJob'>Periodic Job<BR></td></tr>"			
			Response.Write "<tr><td></td><td><font size='2'><input type=checkbox name=RunTime value='Restrictions'>Restrictions<BR></td></tr>"

		Case "After Job"
			Response.Write "<form method='post' action='display.asp' onsubmit='return FrontPage_Form3c_Validator(this)' name='FrontPage_Form3c'><table>"
			Response.Write "<tr><td><font size='2'>Subbmitted By:</td><TD><font size='2'>" & Session("Name")& "<input type =hidden value='" & Session("Name")& "' name=Name></td></tr>"
			Response.Write "<tr><td><font size='2'>Job Name:</td><TD><font size='2'>" & Session("JobName")& "<input type =hidden value='" & Session("JobName")& "' name=JobName></td></tr>"
			Response.Write "<tr><td><font size='2'>Job Class:</td><TD><font size='2'>" & Session("JobClass")& "<input type =hidden value='" & Session("JobClass")& "' name=JobClass></td></tr>"
			Response.Write "<tr><td><font size='2'>Target Host:</td><TD><font size='2'>" & Session("TargetHost")& "<input type =hidden value='" & Session("TargetHost")& "' name=TargetHost></td></tr>"
			Response.Write "<tr><td><font size='2'>Functional Area:</td><TD><font size='2'>" & Session("FunctionalArea")& "<input type =hidden value='" & Session("FunctionalArea")& "' name=FunctionalArea></td></tr>"
			Response.Write "<tr><td><font size='2'>Step Type:</td><TD><font size='2'>" & Session("StepType")& "<input type =hidden value='" & Session("StepType")& "' name=StepType></td></tr>"
	
			Select Case Session("StepType")
				Case "ABAP Program"
					Session("ProgramName1") = Request.Form("ProgramName1")
					Session("VariantName1") = Request.Form("VariantName1")
					Session("ProgramName2") = Request.Form("ProgramName2")
					Session("VariantName2") = Request.Form("VariantName2")
					Session("ProgramName3") = Request.Form("ProgramName3")
					Session("VariantName3") = Request.Form("VariantName3")
					Session("ProgramName4") = Request.Form("ProgramName4")
					Session("VariantName4") = Request.Form("VariantName4")
					Session("ProgramName5") = Request.Form("ProgramName5")
					Session("VariantName5") = Request.Form("VariantName5")

					Session("Language") = Request.Form("Language")
					
					Response.Write "<tr><td><font size='2'>Program Name 1:</td><TD><font size='2'>" & Session("ProgramName1")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 1:</td><TD><font size='2'>" & Session("VariantName1")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 2:</td><TD><font size='2'>" & Session("ProgramName2")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 2:</td><TD><font size='2'>" & Session("VariantName2")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 3:</td><TD><font size='2'>" & Session("ProgramName3")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 3:</td><TD><font size='2'>" & Session("VariantName3")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 4:</td><TD><font size='2'>" & Session("ProgramName4")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 4:</td><TD><font size='2'>" & Session("VariantName4")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 5:</td><TD><font size='2'>" & Session("ProgramName5")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 5:</td><TD><font size='2'>" & Session("VariantName5")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Language:</td><TD><font size='2'>" & Session("Language")& "</td></tr>"
				Case "External Command"
					Session("CommandName") = Request.Form("CommandName")
					Session("Parameters") = Request.Form("Parameters")
					Session("OperatingSystem") = Request.Form("OperatingSystem")
					
					Response.Write "<tr><td><font size='2'>Command Name:</td><TD><font size='2'>" & Session("CommandName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameters:</td><TD><font size='2'>" & Session("Parameters")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Operating System:</td><TD><font size='2'>" & Session("OperatingSystem")& "</td></tr>"
				Case "External Program"
					Session("ProgramName") = Request.Form("ProgramName")
					Session("Parameter") = Request.Form("Parameter")
					
					Response.Write "<tr><td><font size='2'>ProgramName:</td><TD><font size='2'>" & Session("ProgramName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameter:</td><TD><font size='2'>" & Session("Parameter")& "</td></tr>"
			End Select
				Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "<input type=hidden value='" & Session("StartTime")& "' name=StartTime></td></tr>"
				Response.Write "<tr><td><font size='2'>After Job:</td><TD><font size='2'>" & "<input type=text value='" & Session("AfterJob")& "' name=AfterJob></td></tr>"
				Response.Write "<tr><td><font size='2'>Recipient User Name:</td><td><font size='2'><input type=Text name=RecipientName value=></td></tr>"
				Response.Write "<tr><td valign=top><font size='2'>Type of Output Desired:</td><td><font size='2'>"
			' Check boxes for Output Type
				Response.Write "<input type=checkbox name=OutputType value='Copy Required' checked>Copy Required<BR>"
				Response.Write "<input type=checkbox name=OutputType value='Copy'>Copy<BR>"
				Response.Write "<input type=checkbox name=OutputType value='No Printing'>No Printing<BR>"
				Response.Write "<input type=checkbox name=OutputType value='Express'>Express<BR>"
				Response.Write "<input type=checkbox name=OutputType value='To Do'>To Do<BR>"
				Response.Write "<input type=checkbox name=OutputType value='Blind Copy'>Blind Copy<BR>"
				Response.Write "<input type=checkbox name=OutputType value='No Forwarding'>No Forwarding<BR>"
				Response.Write "<input type=checkbox name=OutputType value='Repeat Send'>Repeat Send<BR>"
			
	
		Case "After Event"
			Response.Write "<form method='post' action='index4.asp' onsubmit='return FrontPage_Form3d_Validator(this)' name='FrontPage_Form3d'><table>"
			Response.Write "<tr><td><font size='2'>Subbmitted By:</td><TD><font size='2'>" & Session("Name")& "<input type =hidden value='" & Session("Name")& "' name=Name></td></tr>"
			Response.Write "<tr><td><font size='2'>Job Name:</td><TD><font size='2'>" & Session("JobName")& "<input type =hidden value='" & Session("JobName")& "' name=JobName></td></tr>"
			Response.Write "<tr><td><font size='2'>Job Class:</td><TD><font size='2'>" & Session("JobClass")& "<input type =hidden value='" & Session("JobClass")& "' name=JobClass></td></tr>"
			Response.Write "<tr><td><font size='2'>Target Host:</td><TD><font size='2'>" & Session("TargetHost")& "<input type =hidden value='" & Session("TargetHost")& "' name=TargetHost></td></tr>"
			Response.Write "<tr><td><font size='2'>Functional Area:</td><TD><font size='2'>" & Session("FunctionalArea")& "<input type =hidden value='" & Session("FunctionalArea")& "' name=FunctionalArea></td></tr>"
			Response.Write "<tr><td><font size='2'>Step Type:</td><TD><font size='2'>" & Session("StepType")& "<input type =hidden value='" & Session("StepType")& "' name=StepType></td></tr>"
	
			Select Case Session("StepType")
				Case "ABAP Program"
					Session("ProgramName1") = Request.Form("ProgramName1")
					Session("VariantName1") = Request.Form("VariantName1")
					Session("ProgramName2") = Request.Form("ProgramName2")
					Session("VariantName2") = Request.Form("VariantName2")
					Session("ProgramName3") = Request.Form("ProgramName3")
					Session("VariantName3") = Request.Form("VariantName3")
					Session("ProgramName4") = Request.Form("ProgramName4")
					Session("VariantName4") = Request.Form("VariantName4")
					Session("ProgramName5") = Request.Form("ProgramName5")
					Session("VariantName5") = Request.Form("VariantName5")

					Session("Language") = Request.Form("Language")
					
					Response.Write "<tr><td><font size='2'>Program Name 1:</td><TD><font size='2'>" & Session("ProgramName1")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 1:</td><TD><font size='2'>" & Session("VariantName1")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 2:</td><TD><font size='2'>" & Session("ProgramName2")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 2:</td><TD><font size='2'>" & Session("VariantName2")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 3:</td><TD><font size='2'>" & Session("ProgramName3")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 3:</td><TD><font size='2'>" & Session("VariantName3")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 4:</td><TD><font size='2'>" & Session("ProgramName4")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 4:</td><TD><font size='2'>" & Session("VariantName4")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 5:</td><TD><font size='2'>" & Session("ProgramName5")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 5:</td><TD><font size='2'>" & Session("VariantName5")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Language:</td><TD><font size='2'>" & Session("Language")& "</td></tr>"
				Case "External Command"
					Session("CommandName") = Request.Form("CommandName")
					Session("Parameters") = Request.Form("Parameters")
					Session("OperatingSystem") = Request.Form("OperatingSystem")
					
					Response.Write "<tr><td><font size='2'>Command Name:</td><TD><font size='2'>" & Session("CommandName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameters:</td><TD><font size='2'>" & Session("Parameters")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Operating System:</td><TD><font size='2'>" & Session("OperatingSystem")& "</td></tr>"
				Case "External Program"
					Session("ProgramName") = Request.Form("ProgramName")
					Session("Parameter") = Request.Form("Parameter")
					
					Response.Write "<tr><td><font size='2'>ProgramName:</td><TD><font size='2'>" & Session("ProgramName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameter:</td><TD><font size='2'>" & Session("Parameter")& "</td></tr>"
			End Select
			Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "<input type=hidden value='" & Session("StartTime")& "' name=StartTime></td></tr>"			
			Response.Write "<tr><td><font size='2'>Event Name:</td><TD><font size='2'><input type=name value='" & Session("EventName")& "' name=EventName></td></tr>"
			Response.Write "<tr><td><font size='2'>Parameter:</td><TD><font size='2'><input type=name value='" & Session("Parameter")& "' name=Parameter></td></tr>"
			Response.Write "<tr><TD><font size='2'>Run Time:</td><td><font size='2'><input type=checkbox name=RunTime value='PeriodicJob'>Periodic Job<BR></td></tr>"	
	
		Case "At Operation Mode"
			Response.Write "<form method='post' action='display.asp' onsubmit='return FrontPage_Form3e_Validator(this)' name='FrontPage_Form3e'><table>"
			Response.Write "<tr><td><font size='2'>Subbmitted By:</td><TD><font size='2'>" & Session("Name")& "<input type =hidden value='" & Session("Name")& "' name=Name></td></tr>"
			Response.Write "<tr><td><font size='2'>Job Name:</td><TD><font size='2'>" & Session("JobName")& "<input type =hidden value='" & Session("JobName")& "' name=JobName></td></tr>"
			Response.Write "<tr><td><font size='2'>Job Class:</td><TD><font size='2'>" & Session("JobClass")& "<input type =hidden value='" & Session("JobClass")& "' name=JobClass></td></tr>"
			Response.Write "<tr><td><font size='2'>Target Host:</td><TD><font size='2'>" & Session("TargetHost")& "<input type =hidden value='" & Session("TargetHost")& "' name=TargetHost></td></tr>"
			Response.Write "<tr><td><font size='2'>Functional Area:</td><TD><font size='2'>" & Session("FunctionalArea")& "<input type =hidden value='" & Session("FunctionalArea")& "' name=FunctionalArea></td></tr>"
			Response.Write "<tr><td><font size='2'>Step Type:</td><TD><font size='2'>" & Session("StepType")& "<input type =hidden value='" & Session("StepType")& "' name=StepType></td></tr>"
	
			Select Case Session("StepType")
				Case "ABAP Program"
					Session("ProgramName1") = Request.Form("ProgramName1")
					Session("VariantName1") = Request.Form("VariantName1")
					Session("ProgramName2") = Request.Form("ProgramName2")
					Session("VariantName2") = Request.Form("VariantName2")
					Session("ProgramName3") = Request.Form("ProgramName3")
					Session("VariantName3") = Request.Form("VariantName3")
					Session("ProgramName4") = Request.Form("ProgramName4")
					Session("VariantName4") = Request.Form("VariantName4")
					Session("ProgramName5") = Request.Form("ProgramName5")
					Session("VariantName5") = Request.Form("VariantName5")

					Session("Language") = Request.Form("Language")
					
					Response.Write "<tr><td><font size='2'>Program Name 1:</td><TD><font size='2'>" & Session("ProgramName1")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 1:</td><TD><font size='2'>" & Session("VariantName1")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 2:</td><TD><font size='2'>" & Session("ProgramName2")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 2:</td><TD><font size='2'>" & Session("VariantName2")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 3:</td><TD><font size='2'>" & Session("ProgramName3")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 3:</td><TD><font size='2'>" & Session("VariantName3")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 4:</td><TD><font size='2'>" & Session("ProgramName4")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 4:</td><TD><font size='2'>" & Session("VariantName4")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Program Name 5:</td><TD><font size='2'>" & Session("ProgramName5")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Variant Name 5:</td><TD><font size='2'>" & Session("VariantName5")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Language:</td><TD><font size='2'>" & Session("Language")& "</td></tr>"
				Case "External Command"
					Session("CommandName") = Request.Form("CommandName")
					Session("Parameters") = Request.Form("Parameters")
					Session("OperatingSystem") = Request.Form("OperatingSystem")
					
					Response.Write "<tr><td><font size='2'>Command Name:</td><TD><font size='2'>" & Session("CommandName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameters:</td><TD><font size='2'>" & Session("Parameters")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Operating System:</td><TD><font size='2'>" & Session("OperatingSystem")& "</td></tr>"
				Case "External Program"
					Session("ProgramName") = Request.Form("ProgramName")
					Session("Parameter") = Request.Form("Parameter")
					
					Response.Write "<tr><td><font size='2'>ProgramName:</td><TD><font size='2'>" & Session("ProgramName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameter:</td><TD><font size='2'>" & Session("Parameter")& "</td></tr>"
			End Select
			Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "<input type=hidden value='" & Session("StartTime")& "' name=StartTime></td></tr>"			
			%>
			<!--#include file="OperationMode.asp"-->
			<%

			Response.Write "<tr><td><font size='2'>Recipient User Name:</td><td><font size='2'><input type=Text name=RecipientName value=></td></tr>"
			Response.Write "<tr><td valign=top><font size='2'>Type of Output Desired:</td><td><font size='2'>"
			' Check boxes for Output Type
			Response.Write "<input type=checkbox name=OutputType value='Copy Required' checked>Copy Required<BR>"
			Response.Write "<input type=checkbox name=OutputType value='Copy'>Copy<BR>"
			Response.Write "<input type=checkbox name=OutputType value='No Printing'>No Printing<BR>"
			Response.Write "<input type=checkbox name=OutputType value='Express'>Express<BR>"
			Response.Write "<input type=checkbox name=OutputType value='To Do'>To Do<BR>"
			Response.Write "<input type=checkbox name=OutputType value='Blind Copy'>Blind Copy<BR>"
			Response.Write "<input type=checkbox name=OutputType value='No Forwarding'>No Forwarding<BR>"
			Response.Write "<input type=checkbox name=OutputType value='Repeat Send'>Repeat Send<BR>"
							
	End Select
	
	Response.Write "</td></tr>"
%>
<P>&nbsp;</P>

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

