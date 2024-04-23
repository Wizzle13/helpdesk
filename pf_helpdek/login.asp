<!--
Programer: Chris Burton    
Date Started: 7-9-01

Description:
This Page is the login page for the User interface for the Helpdesk.
This page uses JavaScript to check that there is a user name and 
Password entered.

-->

<html>
<head>
	<title>IT HELP DESK LOGIN</title>
</head>
<link rel="stylesheet" href="helpdesk.css" type="text/css">
<script language="JavaScript">
function EmailFocus(){
<%
If Request.Cookies("Logon") = "" Then 
%>
	Join_Form1.Logon_Name.focus();
<%
Else
%>
	Join_Form1.Password.focus();
<%
End If
%>
}

function Join_Form1_Validator(theForm){	
  if (theForm.Logon_Name.value == "")
  {
    alert("Please enter a value for the \"Logon Name\" field.");
    theForm.Logon_Name.focus();
    return (false);
  }
  
  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";
  var checkStr = theForm.Logon_Name.value;
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
<script language="Javascript">
	
function jumpit(){
	window.location=document.DomainList.Domain.value
	setCookie ("Domain", document.DomainList.Domain.value, now)
	return false
}


</script>
<body onLoad="EmailFocus()">

<form method=post action=_validate.asp onsubmit='return Join_Form1_Validator(this)' onChange='return jumpit()' name='Join_Form1'>
<center><a href='http://purefishing'><img src='/images/pflogo.gif' border=0></a></center><P>&nbsp;</P>
<%
If Request("login") = "n" Then
	Response.Write "<center><font color=Red>Your password is incorrect, please try again.</font></center>"
End If

Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
Set objRs=objConn.Execute("Select DISTINCT Domain from Domains")
 %>
	<table border="0" align=center  width=45%>
	<tr>
		<td valign="bottom">E-mail Address:</td><td valign="bottom"><input type="Text" name="Logon_Name" size="15" tabindex=1 value="<%=Request.Cookies("Logon")%>"><Select name=Domain tabindex=4>
<%
	If Request.Cookies("Domain") = "" Then
		While Not objRs.EOF
			If objRs("Domain") = "@purefishing.com" Then
				Response.Write "<Option selected>"& objRs("Domain") &""
			Else
				Response.Write "<Option>"& objRs("Domain") &""
			End If
			objRs.MoveNext
		Wend
	Else
		While Not objRs.EOF
			If objRs("Domain") = Request.Cookies("Domain") Then
				Response.Write "<Option selected>"& Request.Cookies("Domain") &""
			Else
				Response.Write "<Option>"& objRs("Domain") &""
			End If
			objRs.MoveNext
		Wend
	End If
	objConn.Close
	
	Set objConn = Nothing
 %>		
 </select></td>
	</tr>
	<tr>
		<td>Password:</td><td> <input type="Password" name="Password" size="10" tabindex=2>  <a href="passwordlookup.asp" class="pw">Forget your password?</a></td>
	</tr>
	<tr>
		<td>&nbsp;</td><td><input type="checkbox" name="CookieCheck" tabindex=3
<%
If Request.Cookies("Logon") <> "" Then
%>
checked
<%
End If
%>
>Remember my E-mail Address.
	</tr>
	<TR>
		<td>&nbsp;</td><td><input type="submit" value="Login">&nbsp;<input type="reset" value="Reset">&nbsp;<input type="button" name="join" value="Join" onClick="location.href='join.asp';"></td>
	</tr>
	<TR><Td colspan=2>
	<P>&nbsp;</P>
	<UL>
	<LI>If you are new here, or need help, please read the <a href="faq.asp">FAQs</a>
	</UL>
	</table>
</form>

<center><I>Pure Fishing Online Help Desk v.3.2.17</I></font>
</body></html>
