<!-- This page is displayed as part of the header -->
<script>
<!-- Script to open User selection window.
function remote(){
win2=window.open("chat/index.asp","win2","width=600,height=400,titlebar=0,scrollbars=0, status=0")
win2.creator=self
}
//-->
</script>

<SCRIPT type=text/javascript>
//MENU TITLE
eyesys_title="Online Help Desk"
//TITLE BACKGROUND COLORS
eyesys_titlecol1="black"
eyesys_titlecol2="blue"
//TITLE COLOR
eyesys_titletext="white"
//MENU & ITEM BACKGROUND COLOR
eyesys_bg="#CCCCCC"
//ITEM BACKGROUND COLOR ON MOUSE OVER
eyesys_bgov="blue"
//MENU TEXT COLOR
eyesys_cl="black"
//MENU TEXT COLOR ON MOUSE OVER
eyesys_clov="white"
//MENU WIDTH
eyesys_width=160
//menu starts here
eyesys_init()
//menu item sintax:  eyesys_item(text,icon,link)
//for no icon use 'null'
<%
IF Session("Status")="Manager" or Session("Status")="Worker" or Session("Status")="Administrator" then
%>
eyesys_item('Home','/images/home.gif','/main.asp')
eyesys_item('Add Ticket','/images/add.gif','/add.asp')
eyesys_item('Search Tools','/images/search.gif','/search.asp')
eyesys_item('Maintentance','/images/tools2.gif','/maintenance/')
eyesys_item('Reports','/images/reports.gif','/reports/')
<%
End If
If Session("Status")="Manager" or Session("Status")="Administrator" then
Response.Write "eyesys_item('Manage Users','/images/users.gif','/manageusers.asp')"
End If

If Session("Status")="Member" then
%>
eyesys_item('Home','/images/home.gif','/main.asp')
eyesys_item('Open Ticket','/images/add.gif','/open.asp')
eyesys_item('SAP Job Request','/images/sap.gif','/sapjob/')
eyesys_item('Find Ticket','/images/search.gif','/viewticket.asp')
eyesys_item('Change Password','/images/password.gif','/password.asp')
eyesys_item('HELP!','/images/help.gif','/help.asp')
<%
End If
%>
eyesys_item('Logoff','/images/exit.gif','/logoff.asp')
//mene closes here
eyesys_close()
</SCRIPT>

<table border=0 cellpadding=1 cellspacing=1 width=125 bordercolordark=white bordercolorlight=white bordercolor=white>

<tr><td bgcolor=#8DC2D4 width=6 id=menu_home>
<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_home').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_home').style.backgroundColor = '#8DC2D4';" style='background-color: white;' class = 'hand' onClick=location.href='/main.asp'><a class=menu href="/main.asp">Home</a></td></tr>
<tr><td colspan=2 style='background-color: white;'>&nbsp;</td></tr>
<BR>
<%
' This section displays links that are displayed for user that have the status of Manager or Worker
IF Session("Status")="Manager" or Session("Status")="Worker" or Session("Status")="Administrator" then %>
	<tr><td colspan=2 style='background-color: white;'><B>Admin Tools</B></tr></td>
	<tr><td bgcolor=#8DC2D4 width=6 border=1 id="menu_add">
	<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_add').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_add').style.backgroundColor = '#8DC2D4';" class = 'hand' onClick=location.href='/add.asp'><a class=menu href="/Add.asp">Add Ticket</a></td></tr>
	<tr><td bgcolor=#8DC2D4 width=6 id="menu_sap">
	<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_sap').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_sap').style.backgroundColor = '#8DC2D4';" style='background-color: white;' class = 'hand' onClick=location.href='/saptickets.asp'><a class=menu href="/saptickets.asp">SAP Tickets</a></td></tr>
	<tr><td bgcolor=#8DC2D4 width=6 id=menu_search>
	<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_search').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_search').style.backgroundColor = '#8DC2D4';" style='background-color: white;' class = 'hand' onClick=location.href='/search.asp'><a class=menu href="/search.asp">Search Tools</a></td></tr>
	<tr><td bgcolor=#8DC2D4 width=6 id="menu_maint">
	<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_maint').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_maint').style.backgroundColor = '#8DC2D4';" style='background-color: white;' class = 'hand' onClick=location.href='/maintenance/index.asp'><a class=menu href="/Maintenance/index.asp">Maintenance</a></td></tr>
	<tr><td bgcolor=#8DC2D4 width=6 id="menu_reports">
	<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_reports').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_reports').style.backgroundColor = '#8DC2D4';"  style='background-color: white;' class = 'hand' onClick=location.href='/reports'><a class=menu href="/Reports">Reports</a></td></tr>

<%
' This section displays links that are displayed for user that have the status of Manager
	IF Session("Status")="Manager" or Session("Status")="Administrator" Then %>
	<tr><td bgcolor=#8DC2D4 width=6 id="menu_manuser">
	<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_manuser').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_manuser').style.backgroundColor = '#8DC2D4';"  style='background-color: white;' class = 'hand' onClick=location.href='/manageusers.asp'><a class=menu href="/ManageUsers.asp">Manage Users</a></td></tr><%
	
	End IF
	
	IF Session("Status")="Administrator" Then %>
	
	<tr><td bgcolor=#8DC2D4 width=6 id="menu_junkies">
	<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_junkies').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_junkies').style.backgroundColor = '#8DC2D4';"  style='background-color: white'; class = 'hand' onClick=location.href='/junkies.asp'><a class=menu href="/Junkies.asp">Junkies</a></td></tr>
	
	<tr><td bgcolor=#8DC2D4 width=6 id="menu_whoson">
	<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_whoson').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_whoson').style.backgroundColor = '#8DC2D4';"  style='background-color: white;' class = 'hand' onClick=location.href='/whoson.asp'><a class=menu href="/WhosOn.asp">Who's On?</a></td></tr>
	
	<tr><td bgcolor=#8DC2D4 width=6 id="menu_laston">
	<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_laston').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_laston').style.backgroundColor = '#8DC2D4';"  style='background-color: white;' class = 'hand' onClick=location.href='/laston.asp'><a class=menu href="/LastOn.asp">Last On</a></td></tr>

<%
	End if

If Session("Department")="IT - Systems Team" Then
	Response.Write "<tr><td colspan=2 style='background-color: white;'>&nbsp;</td></tr>"
	Response.Write "<tr><td colspan=2 style='background-color: white;'><P><B>Systems Team</B></td></tr>"
%>
	<tr><td bgcolor=#8DC2D4 width=6 id="menu_secure_links">
	<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_secure_links').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_secure_links').style.backgroundColor = '#8DC2D4';"  style='background-color: white;' class = 'hand' onClick=location.href='/securelinks.asp'><a class=menu href="/securelinks.asp">Secure Links</a></td></tr>

	<tr><td bgcolor=#8DC2D4 width=6 id="menu_call_back">
	<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_call_back').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_call_back').style.backgroundColor = '#8DC2D4';"  style='background-color: white;' class = 'hand' onClick=location.href='/CallBack.asp'><a class=menu href="/CallBack.asp">Call Backs</a></td></tr>

<%
End If
	Response.Write "<tr><td colspan=2 style='background-color: white;'>&nbsp;</td></tr>"
	Response.Write "<tr><td colspan=2 style='background-color: white;'><P><B>User Tools</B></td></tr>"
End If


' This section displays links that are displayed for all Users %>

<tr><td bgcolor=#8DC2D4 width=6 id="menu_open">
<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_open').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_open').style.backgroundColor = '#8DC2D4';"  style='background-color: white;' class = 'hand' onClick=location.href='/open.asp'><a class=menu href="/Open.asp">Open Ticket</a></td></tr>

<tr><td bgcolor=#8DC2D4 width=6 id="menu_sapjob">
<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_sapjob').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_sapjob').style.backgroundColor = '#8DC2D4';" style='background-color: white;' class = 'hand' onClick=location.href='/sapjob/index.asp'><a class=menu href="/Sapjob/index.asp">SAP&nbsp;Job&nbsp;Request</a></td></tr>

<tr><td bgcolor=#8DC2D4 width=6 id="menu_viewticket">
<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_viewticket').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_viewticket').style.backgroundColor = '#8DC2D4';"  style='background-color: white;' class = 'hand' onClick=location.href='/UserSearch.asp'><a class=menu href="/UserSearch.asp">Search Ticket</a></td></tr>

<tr><td bgcolor=#8DC2D4 width=6 id="menu_pref">
<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_pref').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_pref').style.backgroundColor = '#8DC2D4';"  style='background-color: white;' class = 'hand' onClick=location.href='/pref.asp'><a class=menu href="/Pref.asp">User Info</a></td></tr>

<tr><td bgcolor=#8DC2D4 width=6 id="menu_newhire">
<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_newhire').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_newhire').style.backgroundColor = '#8DC2D4';"  style='background-color: white;' class = 'hand' onClick=location.href='/newhire/index.asp'><a class=menu href="/NewHire/index.asp">New Hire</a></td></tr>

<tr><td bgcolor=#8DC2D4 width=6 id="menu_password">
<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_password').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_password').style.backgroundColor = '#8DC2D4';"  style='background-color: white;' class = 'hand' onClick=location.href='/password.asp'><a class=menu href="/password.asp">Change&nbsp;Password</a>&nbsp;</td></tr>

<tr><td bgcolor=#8DC2D4 width=6 id="menu_help">
<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_help').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_help').style.backgroundColor = '#8DC2D4';"  style='background-color: white;' class = 'hand' onClick=location.href='/help.asp'><a class=menu href="/help.asp">HELP!</a></td></tr>

<tr><td bgcolor=#8DC2D4 width=6 id="menu_logoff">
<td onMouseOver="this.style.backgroundColor='#8DC2D4';document.getElementById('menu_logoff').style.backgroundColor = '#015772';" onMouseOut="this.style.backgroundColor='white'; document.getElementById('menu_logoff').style.backgroundColor = '#8DC2D4';"  style='background-color: white;' class = 'hand' onClick=location.href='/logoff.asp'><a class=menu href="/logoff.asp">Logoff</a></td></tr>

</table>
