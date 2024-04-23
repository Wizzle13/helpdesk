<!--#include file="_head.asp"-->
<td valign=top>
<script language="Javascript">
	function checknumber(){
		var x=document.TickNum.Num.value
		var anum=/(^\d+$)|(^\d+\.\d+$)/
		if (anum.test(x))
		testresult=true
	else{
		alert("That's not a number. Please Enter a Number.")
		testresult=false
	}
	return (testresult)
}

	function checkban(){
		if (document.layers||document.all)
			return checknumber()
		else
			return true
}
</script>
<% 
Session("PageName") = "Search"
	Response.Write "<table border=1>"
	

	Response.Write "<tr><td valign=top><form method=post action=SearchWorker.asp?search=Worker>"
	Response.Write"Person Working Call:</td>"
	Response.Write"<td><table border=0><tr><td><select name=ServicedBy>"
		Set objConn=Server.CreateObject("ADODB.Connection")
	objConn.Open "DSN=HelpDesk2"
	Set objRs=objConn.Execute("Select * from Workers where active_worker='Yes' order by First_Name, Last_Name")
	sqlId ="SELECT WORKER_ID FROM Workers;"

	While Not objRs.EOF
	Response.Write"<Option Value="& objRs("Worker_ID") &">"& objRs("First_Name") &" "& objRs("Last_Name")
		objRs.MoveNext
	Wend
	objConn.Close
	Set objConn = Nothing	
	Response.Write "</select></td></tr>"
	Response.Write "<tr><td><input type=radio value =No name=View checked>Open<input type=radio value=Yes name=View>Closed<input type=radio value=All name=View>All"
	Response.Write "<input type=Submit value=Search></form></td></tr></table></td></tr>"

	Response.Write "<tr><td valign=top>Ticket Number:</td>"
	Response.Write "<td><form method=post action=Modifyticket.asp onSubmit='return checkban()' name='TickNum'>"
	Response.Write "<input type='text' name='Num' size='15'>"
	Response.Write "<input type=Submit value=Search></form></td></tr>"
	
	
	Response.Write "<tr><td valign=top>User's First Name:</td>"
	Response.Write "<td><table border=0><tr><td><form method=post action=SearchWorker.asp?search=FName>"
	Response.Write "<input type=Text name=FirstName size=15>"
	Response.Write "<tr><td><input type=radio value =No name=View checked>Open<input type=radio value=Yes name=View>Closed<input type=radio value=All name=View>All"
	Response.Write "<input type=Submit value=Search></form></td></tr></table></td></tr>"

	Response.Write "<tr><td valign=top>User's Last Name:</td>"
	Response.Write "<td><table border=0><tr><td><form method=post action=SearchWorker.asp?search=LName>"
	Response.Write "<input type=Text name=LastName size=15>"
	Response.Write "<tr><td><input type=radio value =No name=View checked>Open<input type=radio value=Yes name=View>Closed<input type=radio value=All name=View>All"
	Response.Write "<input type=Submit value=Search></form></td></tr></table></td></tr>"

	
	Set objConn=Server.CreateObject("ADODB.Connection")
	objConn.Open "DSN=Helpdesk2"
	Set objRs=objConn.Execute("Select * from Domains")
	Response.Write "<tr><td valign=top>Email:</td>"
	Response.Write "<td><table border=0><tr><td><form method=post action=SearchWorker.asp?search=Email>"
	Response.Write "<input type=Text name=Logon_Name size=15><Select name=Country>"
		While Not objRs.EOF
			Response.Write "<Option Value='" & objRs("Country") & "'>"& objRs("Domain") &""
			objRs.MoveNext
		Wend
	objConn.Close
	Set objConn = Nothing		
	
	
	Response.Write "<tr><td><input type=radio value =No name=View checked>Open<input type=radio value=Yes name=View>Closed<input type=radio value=All name=View>All"
	Response.Write "<input type=Submit value=Search></form></td></tr></table></td></tr>"

	Response.Write "<tr><td valign=top><form method=post action=SearchWorker.asp?search=SAP_Module>"
	Response.Write"SAP Module:</td>"
	Response.Write"<td><table border=0><tr><td><select name=SAP_Module>"
	Response.Write"<Option Value=N/A Selected>N/A"
	Response.Write"<Option Value=SD>SD"
	Response.Write"<Option Value=MM>MM"
	Response.Write"<Option Value=PP>PP"
	Response.Write"<Option Value=FI/CO>FI/CO"
	Response.Write"<Option Value=Basis>Basis"
	Response.Write"<Option Value=WM>WM"
	Response.Write "</select></td></tr>"
	Response.Write "<tr><td><input type=radio value =No name=View checked>Open<input type=radio value=Yes name=View>Closed<input type=radio value=All name=View>All"
	Response.Write "<input type=Submit value=Search></form></td></tr></table></td></tr>"
	
	Response.Write"<tr><td valign=top>Package Name:</td><td><table border=0><tr><td><form method=post action=SearchWorker.asp?search=PackageName><select name=PackageName>"
	Set objConn=Server.CreateObject("ADODB.Connection")
	objConn.Open "DSN=HelpDesk2"
	Set objRs=objConn.Execute("Select * from GROUP_MEMBERS order by Package_Name")
	
	While Not objRs.EOF
	Response.Write"<Option value='"& objRs("Package_name")&"'>"& objRs("Package_Name")
		objRs.MoveNext
	Wend
	Response.Write"</select>"
	Response.Write "<tr><td><input type=radio value =No name=View checked>Open<input type=radio value=Yes name=View>Closed<input type=radio value=All name=View>All"
	Response.Write "<input type=Submit value=Search></form></td></tr></table></td></tr>"

	objConn.Close
	
	Set objConn = Nothing

	Response.Write"<tr><td valign=top>Tools:</td><td><table border=0><tr><td><form method=post action=Tickets.asp><select name=Number>"
	Response.Write"<Option value=1>Open Help Tickets"
	Response.Write"<Option value=2>Open Request Tickets"
	Response.Write"<Option value=3>Closed Help Tickets"
	Response.Write"<Option value=4>Close Request Tickets"
	Response.Write"<Option value=5>All Help Tickets"
	Response.Write"<Option value=6>All Request Tickets"
	Response.Write"<Option value=7>All Open Tickets"
	Response.Write"<Option value=8>All Closed Tickets"		

	Response.Write"</select>"
	Response.Write "<input type=Submit value=Search></form></td></tr></table></td></tr>"

 %>

<!--#include file="_Date.asp"-->
<!--#include file="_DateWorker.asp"-->
<!--#include file="_DateDue.asp"-->

<%
'-----> Search By Location<-----'
	Response.Write "<tr><td valign=top>Country or Region:</td>"
	Response.Write "<td align=left><form method=post action=regionsearch.asp name='region'>"
	Response.Write "<table><td valign=top>"
	Response.Write "Asia-Pacific<BR><HR>"
	Response.Write "<input type='checkbox' name='Country' value='AU'>Australia<BR>"
	Response.Write "<input type='checkbox' name='Country' value='JP'>Japan<BR>"
	Response.Write "<input type='checkbox' name='Country' value='NZ'>New Zealand<BR>"
	Response.Write "<input type='checkbox' name='Country' value='TW'>Taiwan"
	Response.Write "</td>"
	Response.Write "<TD valign=top>"
	Response.Write "Europe<BR><HR>"
	Response.Write "<input type='checkbox' name='Country' value='DK'>Denmark<BR>"
	Response.Write "<input type='checkbox' name='Country' value='FI'>Finland<BR>"
	Response.Write "<input type='checkbox' name='Country' value='FR'>France<BR>"
	Response.Write "<input type='checkbox' name='Country' value='DE'>Germany<BR>"
	Response.Write "<input type='checkbox' name='Country' value='IT'>Italy<BR>"
	Response.Write "<input type='checkbox' name='Country' value='NO'>Norway<BR>"
	Response.Write "<input type='checkbox' name='Country' value='SE'>Sweden<BR>"
	Response.Write "<input type='checkbox' name='Country' value='UK'>UK"
	Response.Write "</TD>"
	Response.Write "<TD valign=top>"
	Response.Write "North America<BR><HR>"
	Response.Write "<input type='checkbox' name='Country' value='US'>United States<BR>"
	Response.Write "<input type='checkbox' name='Country' value='CA'>Canada"
	Response.Write "</td></table>"
	Response.Write "<input type=radio value=Open name=CountryView checked>Open"
	Response.Write "<input type=radio value=Closed name=CountryView>Closed"
	Response.Write "<input type=radio value=All name=CountryView>All"
	Response.Write "<input type=Submit value=Search></form></td></tr>"
'----->End Search By Location<-----

'-----> Keyword Search <------'	
	Response.Write "<tr><td valign=top>Keyword or Phrase:</td>"
	Response.Write "<td><form method=post action=keywordsearch.asp>"
	Response.Write "<input type='text' size='15' name='Keyword'><input type=Submit value=Search>"
	Response.Write "<BR><Select name=KeywordSearch>"
	Response.Write "<Option selected value=Exact>The exact phrase entered<option value=All>All of the words entered"
	Response.Write "</form></td></tr>"
'******************************
 %>
</table>
</td>
</tr>
</table>

</body>
</html>
