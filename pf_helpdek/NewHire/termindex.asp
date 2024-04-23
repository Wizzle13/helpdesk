<!--#include file="../_head.asp"-->

<!--
Programer: Chris Burton    
Date Started: 4-22-03

Description:
This Page is for the Pure Fishing Termination of Employment Form

-->

<body>

<%
Session("PageName") = "Pure Fishing Termination of Employment Form"

	Response.Write "<td valign=top>"
%>	


<form method="post" action="termsend.asp" name="Inputform">

<TABLE align=left cellspacing=0>
  <TR>
    <TD></TD>
    <TD colspan=2><font size=3>* Be sure to fill out the form completely.<P></FONT><font size=5>Pure Fishing Termination of Employment Form<BR><HR></font></TD>
  </TR>
  <TR>
    <TD rowspan=20 width=50></TD>
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
    <TD>Department:</TD>
    <TD>
	<select size=1 name="department">
	<option> Administration
	<option> Consumer Service
	<option> Customer Service
	<option> Engineering
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
    <TD>Last Day of Employment:</TD>
    <TD><input type="text" name="enddate" size=8></TD>
  
  </TR>
  <TR>
    <TD>Date to Terminate IT Services:</TD>
    <TD><input type="text" name="ITenddate" size=8></TD>
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


