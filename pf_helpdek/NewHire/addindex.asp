<!--#include file="../_head.asp"-->

<!--
Programer: Chris Burton    
Date Started: 4-22-03

Description:
This Page is for the New Hire Forms

-->

<body>

<%
Session("PageName") = "New/Update Employee Form"

	Response.Write "<td valign=top>"
%>	

<form method="post" action="addsend.asp" name="Inputform">

<TABLE align=left cellspacing=0>
  <TR>
    <TD></TD>
    <TD colspan=2><font size=3>* Be sure to fill out the form completely.<P></FONT><font size=5>Pure Fishing New/Update Employee Form<BR><HR></font></TD>
  </TR>
  <TR>
    <TD rowspan=21 width=50></TD>
    <TD>Date:</TD>
    <TD><input size=8 type=text name=day value=<%=FormatDateTime (Date, vbShortDate)%>></TD>
  </TR>
  <TR>
    <TD>Submitted By:</TD>
    <TD><input type="text" name="submittername" size="20" Value=<%=session("Email") %>></TD>
  </TR>
  <TR>
  	<TD colspan=2><HR><BR>Employee Information:<BR></TD>
  </TR>
  <TR>
    <TD>First Name:</TD>
    <TD><input type="text" name="firstname" size="20"></TD></TR>
  <TR>
    <TD>Middle Initial:</TD>
    <TD><input type="text" name="middleinitial" size="1"></TD></TR>
  <TR>
    <TD>Last Name:</TD>
    <TD><input type="text" name="lastname" size="20"></TD></TR>
 <TR>
    <TD>Employment Status:<BR></TD>
    <TD><input type="radio" name="Package_Name" value="New Employee Setup">New Hire
    <input type="radio" name="Package_Name" value="Employee Information Change">Current Employee</TD></TR>
 <TR>
    <TD>Department:</TD>
    <TD>
	<select size=1 name="department">
	<option> Administration
	<option> Consumer Service
	<option> Customer Service
	<option> Engineering
	<option> FFO
	<option> Finance
	<option> Golf
	<option> Human Resources
	<option> Information Technology
	<option> Marketing
	<option> Product Innovation
	<option> Production - Bait Factory
	<option> Production - Line Factory
	<option> Production - Receiving
	<option> Production - Shipping
	<option> Production - Skilled Trades
	<option> Purchasing
	<option> Sales
	<option> Other - Explain below
	</option>
</select>
	</TD></TR>
  <TR>
    <TD>Title:</TD>
    <TD><input type="text" name="title" size=20></TD>
  </TR>
  <TR>
    <TD>Location:</TD>
    <TD>
	<select size=1 name="location">
	<option> Spirit Lake
	<option> Other - Explain below
	</option>
</select>
	</TD></TR>
  <TR>
    <TD>Supervisor/Manager:</TD>
    <TD><input type="text" name="supervisor" size="20"></TD></TR>
  <TR>
  <TR>
    <TD>Expected Start Date:</TD>
    <TD><input type="text" name="startdate" size=8 value=<%=FormatDateTime (Date, vbShortDate)%>></TD></TR>
  <TR>
    <TD colspan=2><HR><BR>Hardware Requirements:<BR>
	<input type=checkbox name="Hardware" value="Desktop">Desktop Computer<BR>
	<input type=checkbox name="Hardware" value="Laptop">Laptop Computer<BR>
	<input type=checkbox name="Hardware" value="Phone">Phone<BR>
	<input type=checkbox name="Hardware" value="Calling Card">Corporate Calling Card<BR>
	<input type=checkbox name="Hardware" value="Other - See below">Other - Explain below.
	</TD>
  </TR>
  <TR>
    <TD colspan=2><HR><BR>Software/Accounts Required:<BR>*Note: Computers are standardized to include<BR>Microsoft Office (Word, Excel, Power Point,<BR>and Access), Internet Explorer, Acrobat Reader, <BR>and Norton Anti-Virus.<P>
Accounts:<BR>
	<input type=checkbox name="Accounts" value="NT">Windows NT<BR>
	<input type=checkbox name="Accounts" value="SAP R/3">SAP R/3 (Standard SAP System)<BR>
	<input type=checkbox name="Accounts" value="SAP APO">SAP APO (Advanced Planning and Optimization)<BR>
	<input type=checkbox name="Accounts" value="SAP BW">SAP BW  (Business Warehouse)<BR>
	
	<input type=checkbox name="Accounts" value="E-mail">E-mail<BR>
	<input type=checkbox name="Accounts" value="OWA">Outlook Web Access<BR>
	<input type=checkbox name="Accounts" value="DirectFax">Direct Fax<BR>
	<input type=checkbox name="Accounts" value="Voicemail">Voicemail<BR>
	<input type=checkbox name="Accounts" value="Text2Speech">Text2Speech<P></TD>
  </tr>
  <TR>
  	<TD>
	Software:<BR>	
	
	<input type=checkbox name="Software" value="AutoCAD">AutoCAD<BR>
	<input type=checkbox name="Software" value="AutoCAD LT">AutoCAD LT<BR>
	<input type=checkbox name="Software" value="DejaVu">DejaVu<BR>
	<input type=checkbox name="Software" value="Demand Solutions">Demand Solutions<BR>
    </TD>
	<TD>
       	<BR>
	<input type=checkbox name="Software" value="HR Perspective">HR Perspective<BR>
	<input type=checkbox name="Software" value="Pro-Engineer">Pro-Engineer<BR>
	<input type=checkbox name="Software" value="Pro-Mechanica">Pro-Mechanica<BR>
	<input type=checkbox name="Software" value="Other - See below">Other - Explain below.<BR>
	</TD>
  </TR>
  <TR>
    	<TD colspan=2>
	<HR>
	<BR>NT Groups: (M:\ Drive Groups)
	<BR>
	<textarea cols=35 rows=10 name="ntgrp">
	</textarea>
	</TD>
  </TR>
  <TR>
	<TD colspan=2>
	<HR>
	<BR>E-mail Distribution Lists:
	<BR>
	<textarea cols=35 rows=10 name="emailgrp">
	</textarea>
	</TD>
  </TR>
  <TR>
	<TD colspan=2>
	<HR>
	<BR>SAP Authorizations:
	<BR>
	<textarea cols=35 rows=10 name="sapauth">
	</textarea>
	</TD>
  </TR>
  <TR>
	<TD colspan=2>
	<HR>
	<BR>Additional Notes:
	<BR>
	<textarea cols=35 rows=10 name="notes">
	</textarea>
	</TD>
  </TR>
  <TR>
    <TD align=right><input type="submit" value="Send" name="B1"></TD>
    <TD align=left><input type="reset"value="Reset" name="B2"></TD>
  </TR>
</TABLE>
</form>

<%
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

%>
</body>
</html>


