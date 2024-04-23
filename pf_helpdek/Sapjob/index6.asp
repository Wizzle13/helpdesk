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
function FrontPage_Form6_Validator(theForm)
{
  if (theForm.TimeSpan.value == "")
  {
    alert("Please enter a value for the \"Time Span\" field.");
    theForm.TimeSpan.focus();
    return (false);
  }


  if (theForm.RecipientName.value == "")
  {
    alert("Please enter a value for the \"Recipient Name\" field.");
    theForm.RecipientName.focus();
    return (false);
  }


  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.TimeSpan.value;
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
</script>


<P>&nbsp;</P>
<font face="Arial" size="4">SAP Background Job Form</font>
<%

			Response.Write "<form method='post' action='display.asp' onsubmit='return FrontPage_Form6_Validator(this)' name='FrontPage_Form6'><table>"
			Response.Write"<input type =Hidden value='0' name=TimeSpan>"				
			Response.Write "<tr><td><font size='2'>Subbmitted By:</td><TD><font size='2'>" & Session("Name")& "<input type =hidden value='" & Session("Name")& "' name=Name></td></tr>"
			Response.Write "<tr><td><font size='2'>Job Name:</td><TD><font size='2'>" & Session("JobName")& "<input type =hidden value='" & Session("JobName")& "' name=JobName></td></tr>"
			Response.Write "<tr><td><font size='2'>Job Class:</td><TD><font size='2'>" & Session("JobClass")& "<input type =hidden value='" & Session("JobClass")& "' name=JobClass></td></tr>"
			Response.Write "<tr><td><font size='2'>Target Host:</td><TD><font size='2'>" & Session("TargetHost")& "<input type =hidden value='" & Session("TargetHost")& "' name=TargetHost></td></tr>"
			Response.Write "<tr><td><font size='2'>Functional Area:</td><TD><font size='2'>" & Session("FunctionalArea")& "<input type =hidden value='" & Session("FunctionalArea")& "' name=FunctionalArea></td></tr>"
			Response.Write "<tr><td><font size='2'>Step Type:</td><TD><font size='2'>" & Session("StepType")& "<input type =hidden value='" & Session("StepType")& "' name=StepType></td></tr>"
	
			Select Case Session("StepType")
				Case "ABAP Program"
				 				
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
									
					Response.Write "<tr><td><font size='2'>Command Name:</td><TD><font size='2'>" & Session("CommandName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameters:</td><TD><font size='2'>" & Session("Parameters")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Operating System:</td><TD><font size='2'>" & Session("OperatingSystem")& "</td></tr>"
				
				Case "External Program"
					
					Response.Write "<tr><td><font size='2'>ProgramName:</td><TD><font size='2'>" & Session("ProgramName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameter:</td><TD><font size='2'>" & Session("Parameter")& "</td></tr>"
			End Select
			Select Case Session("StartTime")
		'******************************************
		'* This is the code to execute if         *
		'* StartTime is "Immediate"               *
		'******************************************	
				
		Case "Date/Time"
			Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "<input type=hidden value='" & Session("StartTime")& "' name=StartTime></td></tr>"
			Response.Write "<tr><td><font size='2'>Date Needed:</td><TD><font size='2'>" & Session("StartHour")& "<input type=hidden value='" & Session("StartHour")& "' name=StartHour> <font size='2'>" & Session("StartMonth")& "<input type=hidden value='" & Session("StartMonth")& "' name=StartMonth> <font size='2'>" & Session("StartDate")& "<input type=hidden value='" & Session("StartDate")& "' name=StartDate></td></tr>"
			Response.Write "<tr><td><font size='2'>No Start After:</td><TD><font size='2'>" & Session("NoStartHour")& "<input type=hidden value='" & Session("NoStartHour")& "' name=NoStartHour> <font size='2'>" & Session("NoStartMonth")& "<input type=hidden value='" & Session("NoStartMonth")& "' name=NoStartMonth> <font size='2'>" & Session("NoStartDate")& "<input type=hidden value='" & Session("NoStartDate")& "' name=NoStartDate></td></tr>"
			Response.Write "<tr><td><font size='2'>Run Time:</td><TD><font size='2'>" & Session("RunTime")& "<input type =hidden value='" & Session("RunTime")& "' name=RunTime></td></tr>"

		Case "After Event"
			Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "<input type=hidden value='" & Session("StartTime")& "' name=StartTime></td></tr>"			
			Session("EventName") = Request.Form("EventName")
			Response.Write "<tr><td><font size='2'>Event Name:</td><TD><font size='2'>" & Session("EventName")& "<input type=hidden value='" & Session("EventName")& "' name=EventName></td></tr>"
			Session("Parameter") = Request.Form("Parameter")
			Response.Write "<tr><td><font size='2'>Parameter:</td><TD><font size='2'>" & Session("Parameter")& "<input type=Hidden value='" & Session("Parameter")& "' name=Parameter></td></tr>"
			Session("RunTime") = Request.Form("RunTime")
			Response.Write "<tr><td><font size='2'>Run Time:</td><TD><font size='2'>" & Session("RunTime")& "<input type =hidden value='" & Session("RunTime")& "' name=RunTime></td></tr>"
End Select			

			Select Case Session("RunTime")
				Case "PeriodicJob"
				Session("PeriodicValues") = Request.Form("PeriodicValues")
				Response.Write "<tr><td><font size='2'>Periodic Values:</td><TD><font size='2'>" & Session("PeriodicValues")& "<input type =hidden value='" & Session("PeriodicValues")& "' name=PeriodicValues></td></tr>"			
				Select Case Session("PeriodicValues")
					Case "Other Period"
						Session("OtherPeriod") = Request.Form("OtherPeriod")				
						Response.Write "<tr><td><font size='2'>Other Period:</td><TD><font size='2'>" & Session("OtherPeriod")& "<input type =hidden value='" & Session("OtherPeriod")& "' name=OtherPeriod></td></tr>"							
						Response.Write "<tr><td><font size='2'> "& Session("OtherPeriod")& "<input type =hidden value='" & Session("OtherPeriod")& "' name=OtherPeriod>:</td><TD><font size='2'>" & "<input type =name value='" & Session("TimeSpan")& "' name=TimeSpan></td></tr>"							
				End Select						
			End Select		
			Select Case Session("RunTime")
				Case "PeriodicJob, Restrictions"
				Session("PeriodicValues") = Request.Form("PeriodicValues")
				Response.Write "<tr><td><font size='2'>Periodic Values:</td><TD><font size='2'>" & Session("PeriodicValues")& "<input type =hidden value='" & Session("PeriodicValues")& "' name=PeriodicValues></td></tr>"			

				Select Case Session("PeriodicValues")
					Case "Other Period"
						Session("OtherPeriod") = Request.Form("OtherPeriod")				
						Response.Write "<tr><td><font size='2'>Other Period:</td><TD><font size='2'>" & Session("OtherPeriod")& "<input type =hidden value='" & Session("OtherPeriod")& "' name=OtherPeriod></td></tr>"							
						Response.Write "<tr><td><font size='2'> "& Session("OtherPeriod")& "<input type =hidden value='" & Session("OtherPeriod")& "' name=OtherPeriod>:</td><TD><font size='2'>" & "<input type =name value='" & Session("TimeSpan")& "' name=TimeSpan></td></tr>"													
				End Select						
				
				Response.Write "<tr><td><font size='2'>Restrictions:</td><TD><font size='2'>" & Session("Restrictions") & "<input type=hidden value='" & Session("Restrictions") & "' name=Execute Only on Weekdays></td></tr>"		
				Session("BehaviorRestriction") = Request.Form("BehaviorRestriction")
				Response.Write "<tr><td><font size='2'>Behavior Restriction:</td><TD><font size='2'>" & Session("BehaviorRestriction") & "<input type=hidden value='" & Session("BehaviorRestriction") & "' name=Behavior Restriction></td></tr>"		
				Session("Factory Calander ID") = Request.Form("Factory Calander ID")
				Response.Write "<tr><td><font size='2'>Factory Calander ID:</td><TD><font size='2'>" & Session("Factory Calander ID") & "<input type=hidden value='" & Session("Factory Calander ID") & "' name=Factory Calander ID></td></tr>"		

			End Select		
			
			Select Case Session("RunTime")
				Case "Restrictions"
				Session("Restrictions") = Request.Form("Restrictions")
				Response.Write "<tr><td><font size='2'>Restrictions:</td><TD><font size='2'>" & Session("Restrictions") & "<input type=hidden value='" & Session("Restrictions") & "' name=Execute Only on Weekdays></td></tr>"		
				Session("BehaviorRestriction") = Request.Form("BehaviorRestriction")
				Response.Write "<tr><td><font size='2'>Behavior Restriction:</td><TD><font size='2'>" & Session("BehaviorRestriction") & "<input type=hidden value='" & Session("BehaviorRestriction") & "' name=Behavior Restriction></td></tr>"		
				Session("Factory Calander ID") = Request.Form("Factory Calander ID")
				Response.Write "<tr><td><font size='2'>Factory Calander ID:</td><TD><font size='2'>" & Session("Factory Calander ID") & "<input type=hidden value='" & Session("Factory Calander ID") & "' name=Factory Calander ID></td></tr>"		

			End Select		
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
				Response.Write "<tr><td><font size='2'>Additional Comments:</td><td><font size='2'><Textarea cols=25 Rows=3 name=Comments value=></textarea></td></tr>"
			
			



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

