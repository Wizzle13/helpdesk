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
		'* StartTime is "Date/Time"               *
		'******************************************	
				
		Case "Date/Time"
			Response.Write "<form method='post' action='index5.asp'><table>"
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
			
			Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "<input type=hidden value='" & Session("StartTime")& "' name=StartTime></td></tr>"
			Session("StartHour") = Request.Form("StartHour")
			Session("StartDate") = Request.Form("StartDate")
			Session("StartMonth") = Request.Form("StartMonth")
			Response.Write "<tr><td><font size='2'>Date Needed:</td><TD><font size='2'>" & Session("StartHour")& "<input type=hidden value='" & Session("StartHour")& "' name=StartHour> <font size='2'>" & Session("StartMonth")& "<input type=hidden value='" & Session("StartMonth")& "' name=StartMonth> <font size='2'>" & Session("StartDate")& "<input type=hidden value='" & Session("StartDate")& "' name=StartDate></td></tr>"
			Session("NoStartHour") = Request.Form("NoStartHour")
			Session("NoStartDate") = Request.Form("NoStartDate")
			Session("NoStartMonth") = Request.Form("NoStartMonth")
			Response.Write "<tr><td><font size='2'>No Start After:</td><TD><font size='2'>" & Session("NoStartHour")& "<input type=hidden value='" & Session("NoStartHour")& "' name=NoStartHour> <font size='2'>" & Session("NoStartMonth")& "<input type=hidden value='" & Session("NoStartMonth")& "' name=NoStartMonth> <font size='2'>" & Session("NoStartDate")& "<input type=hidden value='" & Session("NoStartDate")& "' name=NoStartDate></td></tr>"
			Session("RunTime") = Request.Form("RunTime")
			Response.Write "<tr><td><font size='2'>Run Time:</td><TD><font size='2'>" & Session("RunTime")& "<input type =hidden value='" & Session("RunTime")& "' name=RunTime></td></tr>"
			
			Select Case Session("RunTime")
				Case "PeriodicJob"
					%>		
					<!--#Include file="PerodicJob.asp"-->
					<%	

				Case "Restrictions"
					Response.Write "<tr><td><font size='2'>Restrictions:</td><td><font size='2'><input type=checkbox name=Restrictions value='Only Workdays'>Execute Only on Workdays<BR></td></tr>"
		
				Case "PeriodicJob, Restrictions"
					%>		
					<!--#Include file="PerodicJob.asp"-->
					<%	
					Response.Write "<tr><td><font size='2'>Restrictions:</td><td><font size='2'><input type=checkbox name=Restrictions value='Only Workdays'>Execute Only on Workdays<BR></td></tr>"
			End Select
	
		Case "After Event"
			Response.Write "<form method='post' action='index5.asp'><table>"
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
			Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "<input type=hidden value='" & Session("StartTime")& "' name=StartTime></td></tr>"			
			Session("EventName") = Request.Form("EventName")
			Response.Write "<tr><td><font size='2'>Event Name:</td><TD><font size='2'>" & Session("EventName")& "<input type=hidden value='" & Session("EventName")& "' name=EventName></td></tr>"
			Session("Parameter") = Request.Form("Parameter")
			Response.Write "<tr><td><font size='2'>Parameter:</td><TD><font size='2'>" & Session("Parameter")& "<input type=Hidden value='" & Session("Parameter")& "' name=Parameter></td></tr>"
			Session("RunTime") = Request.Form("RunTime")
			Response.Write "<tr><td><font size='2'>Run Time:</td><TD><font size='2'>" & Session("RunTime")& "<input type =hidden value='" & Session("RunTime")& "' name=RunTime></td></tr>"

			Select Case Session("RunTime")
				Case "PeriodicJob"
					%>		
					<!--#Include file="PerodicJob.asp"-->
					<%	
			End Select					
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

