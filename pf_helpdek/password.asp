<!--
Programer: Chris Burton    
Date Started: 7-9-01

Description:
This page Allows a user to change there password.
This page uses JavaScript to check that all the information is entered.

-->

<script language="JavaScript">
function Join_Form1_ValIDator(theForm)
{	
  if (theForm.OldPassword.value == "")
  {
    alert("Please enter a value for the \"Old Password\" field.");
    theForm.OldPassword.focus();
    return (false);
  }

  if (theForm.NewPassword.value == "")
  {
    alert("Please enter a value for the \"New Password\" field.");
    theForm.NewPassword.focus();
    return (false);
  }
   if (theForm.NewPassword2.value == "")
  {
    alert("Please enter a value for the \"Confirm New Password\" field.");
    theForm.NewPassword2.focus();
    return (false);
  }
   if (theForm.NewPassword.value != theForm.NewPassword2.value)
  {
    alert("Passwords do not match. Please retype passwords.");
    theForm.NewPassword.focus();
    return (false);
  }
 
  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890_";

  var checkStr = theForm.OldPassword.value;
  var checkStr = theForm.NewPassword.value;
  var checkStr = theForm.NewPassword2.value;
  var allValID = true;
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

<!--#include file="_head.asp"-->

<%
	Dim sqlID
	Dim objConn
	Dim objRec
	Dim rsID
	Dim rsName

	'Set objConn = Server.CreateObject ("ADODB.Connection")
	'Set objRec = Server.CreateObject ("ADODB.Recordset")

	
	'objConn.Open "intranet"
	''This Section selects pulls the user info.
	'sqlID ="SELECT * FROM UserInfo WHERE ID=" & Session("UID") & ";"

	'objRec.Open sqlID, objConn	

	Response.Write "<td valign=top>"
	Response.Write "<form method=post action=""Change.asp"" onsubmit='return Join_Form1_ValIDator(this)' name='Join_Form1'>"
	Response.Write "<table border=0>"
	'Response.Write "<td><P>" & Session("FirstName") &" " & Session("LastName") &"</td>"
	'Response.Write "<td>&nbsp;</td></TR>"
	'Response.Write "<tr><td>Old Password:</td><td><input type=text name=Num2 size=20 value='" & Session("password")&"'></td></tr>"
	Response.Write "<tr><td>Old Password:</td><td><input type=Password name=OldPassword size=20></td></tr>"
	Response.Write "<tr><td>New Password:</td><td><input type=Password name=NewPassword size=20></td></tr>"
	Response.Write "<tr><td>Confirm New Password:</td><td><input type=Password name=NewPassword2 size=20></td></tr>"

	Response.Write "<tr><td colspan=2><input type=Submit value=Change></td></tr></TABLE></form>"

	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	
	'objRec.Close
	'objConn.Close
	'Set objRec = Nothing
	'Set objConn = Nothing	
	
%>
<p>






</BODY>
</HTML>

