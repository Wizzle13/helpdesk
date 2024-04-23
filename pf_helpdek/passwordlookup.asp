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
	<title>O.H.D. Password Lookup</title>
</head>
<link rel="stylesheet" href="helpdesk.css" type="text/css">
<script language="JavaScript">
function EmailFocus(){
	Join_Form1.Logon_Name.focus();
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
</script>
<form method=post action=_validate.asp onsubmit='return Join_Form1_Validator(this)' onChange='return jumpit()' name='Join_Form1'>
<center><a href='http://purefishing'><img src='/images/pflogo.gif' border=0></a><P>&nbsp;</P>
<b>O.H.D. Password Lookup</b><br>Please enter your E-mail address below and click "Send Me My Password".<BR>Your password will be E-mailed to you.</center><P>
<%
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"
Set objRs=objConn.Execute("Select DISTINCT Domain from Domains")
 %>
	<table border="0" align=center  width=45%>
	<tr>
		<td valign="bottom">E-mail Address:</td><td valign="bottom"><input type="Text" name="Logon_Name" size="15" tabindex=1><Select name=Domain tabindex=3>
<%
		While Not objRs.EOF
			If objRs("Domain") = "@purefishing.com" Then
				Response.Write "<Option selected>"& objRs("Domain") &""
			Else
				Response.Write "<Option>"& objRs("Domain") &""
			End If
			objRs.MoveNext
	Wend
	objConn.Close
	
	Set objConn = Nothing
 %>		
 </select></td>
	</tr>
	<TR>
		<td>&nbsp;</td><td><input type="submit" value="Send Me My Password"></td>
	</tr>
	<TR><Td colspan=2>
	<P>&nbsp;</P>
	
	</table>
</form>

<center><I>Pure Fishing Online Help Desk v.3.1</I></font>
</body></html>
