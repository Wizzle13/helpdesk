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
			Response.Write "<form method='post' action='send.asp'><table>"
			Response.Write "<tr><td><font size='2'>Submitted By:</td><TD><font size='2'>" & Session("Name")& "</td></tr>"
			Response.Write "<tr><td><font size='2'>Job Name:</td><TD><font size='2'>" & Session("JobName")& "</td></tr>"
			Response.Write "<tr><td><font size='2'>Job Class:</td><TD><font size='2'>" & Session("JobClass")& "</td></tr>"
			Response.Write "<tr><td><font size='2'>Target Host:</td><TD><font size='2'>" & Session("TargetHost")& "</td></tr>"
			Response.Write "<tr><td><font size='2'>Functional Area:</td><TD><font size='2'>" & Session("FunctionalArea")& "</td></tr>"
			Response.Write "<tr><td><font size='2'>Step Type:</td><TD><font size='2'>" & Session("StepType")& "</td></tr>"
	
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
				Case "External Command"
					
					
					Response.Write "<tr><td><font size='2'>Command Name:</td><TD><font size='2'>" & Session("CommandName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameters:</td><TD><font size='2'>" & Session("Parameters")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Operating System:</td><TD><font size='2'>" & Session("OperatingSystem")& "</td></tr>"
				Case "External Program"
					
					
					Response.Write "<tr><td><font size='2'>ProgramName:</td><TD><font size='2'>" & Session("ProgramName")& "</td></tr>"
					Response.Write "<tr><td><font size='2'>Parameter:</td><TD><font size='2'>" & Session("Parameter")& "</td></tr>"
			End Select
			Session("RecipientName") = Request.Form("RecipientName")
			Session("OutputType") = Request.Form("OutputType")
		Select Case Session("StartTime")
		'******************************************
		'* This is the code to execute if         *
		'* StartTime is "Immediate"               *
		'******************************************	
		Case "Immediate"
	
			Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "</td></tr>"
			
			
					
		Case "Date/Time"
  			Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "<input type=hidden value='" & Session("StartTime")& "' name=StartTime></td></tr>"
			Response.Write "<tr><td><font size='2'>Date Needed:</td><TD><font size='2'>" & Session("StartHour")& "<input type=hidden value='" & Session("StartHour")& "' name=StartHour> <font size='2'>" & Session("StartMonth")& "<input type=hidden value='" & Session("StartMonth")& "' name=StartMonth> <font size='2'>" & Session("StartDate")& "<input type=hidden value='" & Session("StartDate")& "' name=StartDate></td></tr>"			
			Response.Write "<tr><td><font size='2'>No Start After:</td><TD><font size='2'>" & Session("NoStartHour")& "<input type=hidden value='" & Session("NoStartHour")& "' name=NoStartHour> <font size='2'>" & Session("NoStartMonth")& "<input type=hidden value='" & Session("NoStartMonth")& "' name=NoStartMonth> <font size='2'>" & Session("NoStartDate")& "<input type=hidden value='" & Session("NoStartDate")& "' name=NoStartDate></td></tr>"			
			Response.Write "<tr><td><font size='2'>Run Time:</td><TD><font size='2'>" & Session("RunTime")& "<input type =hidden value='" & Session("RunTime")& "' name=RunTime></td></tr>"

			Select Case Session("RunTime")
				Case "PeriodicJob"
				Session("PeriodicValues") = Request.Form("PeriodicValues")
				Response.Write "<tr><td><font size='2'>Periodic Values:</td><TD><font size='2'>" & Session("PeriodicValues")& "<input type =hidden value='" & Session("PeriodicValues")& "' name=PeriodicValues></td></tr>"			
				Select Case Session("PeriodicValues")
					Case "Other Period"
						Response.Write "<tr><td><font size='2'>Other Period:</td><TD><font size='2'>" & Session("OtherPeriod")& "<input type =hidden value='" & Session("OtherPeriod")& "' name=OtherPeriod></td></tr>"							
						Session("TimeSpan") = Request.Form("TimeSpan")
						Response.Write "<tr><td><font size='2'> "& Session("OtherPeriod")& "<input type =hidden value='" & Session("OtherPeriod")& "' name=OtherPeriod>:</td><TD><font size='2'>"& Session("TimeSpan")& "<input type =hidden value='" & Session("TimeSpan")& "' name=TimeSpan></td></tr>"							
				End Select						
			End Select		
			Select Case Session("RunTime")
				Case "PeriodicJob, Restrictions"
				Response.Write "<tr><td><font size='2'>Periodic Values:</td><TD><font size='2'>" & Session("PeriodicValues")& "<input type =hidden value='" & Session("PeriodicValues")& "' name=PeriodicValues></td></tr>"			
				Select Case Session("PeriodicValues")
					Case "Other Period"
						Response.Write "<tr><td><font size='2'>Other Period:</td><TD><font size='2'>" & Session("OtherPeriod")& "<input type =hidden value='" & Session("OtherPeriod")& "' name=OtherPeriod></td></tr>"							
						Session("TimeSpan") = Request.Form("TimeSpan")
						Response.Write "<tr><td><font size='2'> "& Session("OtherPeriod")& "<input type =hidden value='" & Session("OtherPeriod")& "' name=OtherPeriod>:</td><TD><font size='2'>"& Session("TimeSpan")& "<input type =hidden value='" & Session("TimeSpan")& "' name=TimeSpan></td></tr>"							

				End Select						
				Response.Write "<tr><td><font size='2'>Restrictions:</td><TD><font size='2'>" & Session("Restrictions") & "<input type=hidden value='" & Session("Restrictions") & "' name=Execute Only on Weekdays></td></tr>"		
				Response.Write "<tr><td><font size='2'>Behavior Restriction:</td><TD><font size='2'>" & Session("BehaviorRestriction") & "<input type=hidden value='" & Session("BehaviorRestriction") & "' name=Behavior Restriction></td></tr>"		
				Response.Write "<tr><td><font size='2'>Factory Calander ID:</td><TD><font size='2'>" & Session("Factory Calander ID") & "<input type=hidden value='" & Session("Factory Calander ID") & "' name=Factory Calander ID></td></tr>"						
			End Select		
			Select Case Session("RunTime")
				Case "Restrictions"
				Response.Write "<tr><td><font size='2'>Restrictions:</td><TD><font size='2'>" & Session("Restrictions") & "<input type=hidden value='" & Session("Restrictions") & "' name=Execute Only on Weekdays></td></tr>"		
				Response.Write "<tr><td><font size='2'>Behavior Restriction:</td><TD><font size='2'>" & Session("BehaviorRestriction") & "<input type=hidden value='" & Session("BehaviorRestriction") & "' name=Behavior Restriction></td></tr>"		
				Response.Write "<tr><td><font size='2'>Factory Calander ID:</td><TD><font size='2'>" & Session("Factory Calander ID") & "<input type=hidden value='" & Session("Factory Calander ID") & "' name=Factory Calander ID></td></tr>"		
			End Select		
			

		Case "After Job"
				Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "<input type=hidden value='" & Session("StartTime")& "' name=StartTime></td></tr>"
				Session("AfterJob") = Request.Form("AfterJob")
				Response.Write "<tr><td><font size='2'>After Job:</td><TD><font size='2'>" &  Session("AfterJob")& "<input type=hidden value='" & Session("AfterJob")& "' name=AfterJob></td></tr>"
				
		Case "After Event"
			Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "<input type=hidden value='" & Session("StartTime")& "' name=StartTime></td></tr>"			
			
			Response.Write "<tr><td><font size='2'>Event Name:</td><TD><font size='2'>" & Session("EventName")& "<input type=hidden value='" & Session("EventName")& "' name=EventName></td></tr>"
			
			Response.Write "<tr><td><font size='2'>Parameter:</td><TD><font size='2'>" & Session("Parameter")& "<input type=Hidden value='" & Session("Parameter")& "' name=Parameter></td></tr>"
			
			Response.Write "<tr><td><font size='2'>Run Time:</td><TD><font size='2'>" & Session("RunTime")& "<input type =hidden value='" & Session("RunTime")& "' name=RunTime></td></tr>"
			Select Case Session("RunTime")
				Case "PeriodicJob"
			
				Response.Write "<tr><td><font size='2'>Periodic Values:</td><TD><font size='2'>" & Session("PeriodicValues")& "<input type =hidden value='" & Session("PeriodicValues")& "' name=PeriodicValues></td></tr>"			
				Select Case Session("PeriodicValues")
					Case "Other Period"
			
						Response.Write "<tr><td><font size='2'>Other Period:</td><TD><font size='2'>" & Session("OtherPeriod")& "<input type =hidden value='" & Session("OtherPeriod")& "' name=OtherPeriod></td></tr>"							
						Session("TimeSpan") = Request.Form("TimeSpan")
						Response.Write "<tr><td><font size='2'> "& Session("OtherPeriod")& "<input type =hidden value='" & Session("OtherPeriod")& "' name=OtherPeriod>:</td><TD><font size='2'>"& Session("TimeSpan")& "<input type =hidden value='" & Session("TimeSpan")& "' name=TimeSpan></td></tr>"							
				End Select						
			End Select		
			Select Case Session("RunTime")
				Case "PeriodicJob, Restrictions"
			
				Response.Write "<tr><td><font size='2'>Periodic Values:</td><TD><font size='2'>" & Session("PeriodicValues")& "<input type =hidden value='" & Session("PeriodicValues")& "' name=PeriodicValues></td></tr>"			

				Select Case Session("PeriodicValues")
					Case "Other Period"

						Response.Write "<tr><td><font size='2'>Other Period:</td><TD><font size='2'>" & Session("OtherPeriod")& "<input type =hidden value='" & Session("OtherPeriod")& "' name=OtherPeriod></td></tr>"							
				End Select						

				Response.Write "<tr><td><font size='2'>Restrictions:</td><TD><font size='2'>" & Session("Execute Only on Weekdays") & "<input type=hidden value='" & Session("Execute Only on Weekdays") & "' name=Execute Only on Weekdays></td></tr>"		

			End Select		
			
			Select Case Session("RunTime")
				Case "Restrictions"

				Response.Write "<tr><td><font size='2'>Restrictions:</td><TD><font size='2'>" & Session("Execute Only on Weekdays") & "<input type=hidden value='" & Session("Execute Only on Weekdays") & "' name=Execute Only on Weekdays></td></tr>"		
			End Select		
		
		Case "At Operation Mode"
				Response.Write "<tr><td><font size='2'>Start Time:</td><TD><font size='2'>" & Session("StartTime")& "</td></tr>"
				Session("ModeName") = Request.Form("ModeName")
				Response.Write "<tr><td><font size='2'> Mode Name: </td><TD><font size='2'>"& Session("ModeName")& "<input type =hidden value='" & Session("ModeName")& "' name=ModeName></td></tr>"							
			End Select
			Response.Write "<tr><td><font size='2'>Recipient User Name:</td><td><font size='2'>" & Session("RecipientName") & "</td></tr>"
			Response.Write "<tr><td valign=top><font size='2'>Type of Output Desired:</td><td><font size='2'>" & Session("OutputType") & "</td></tr>"
			Session("Comments") = Request.Form("Comments")
			Response.Write "<tr><td valign=top><font size='2'>Additional Comments:</td><td><font size='2'>" & Session("Comments") & "</td></tr>"
	Response.Write "</td></tr>"
%>
<P>&nbsp;</P>

		</td>
	</tr>
</table>
<input type="button" onclick="javascript:history.go(-1)" value="<<Back">
<input type="submit" value="Send >>">

</form>

</TD>
</TR>
</TABLE>
</BODY>
</HTML>
