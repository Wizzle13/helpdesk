<%
'Programmer: Chris Burton    
'Date Started: 7-9-01

'Description:
'This page Checks to see if Old Password is the same as the Password. If the Password is the same it then writes the new password in to the database, else it will make the user go back to the previous page to enter the Passwords again.

	Dim objConn
	Dim objRec
	Dim sql
	Dim sqlverify

	
	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")
	
	objConn.Open "Helpdesk2"
	'objRec.Open "UserInfo", objConn
	'This section is supposed to pull the password for the user.
	sqlverify = "SELECT Password FROM UserInfo WHERE ID=" & Session("UID") & ";"
	objRec.Open sqlverify, objConn
		
	sql = "UPDATE UserInfo SET  Password= '" & Request.Form("NewPassword") & "', UserAdded='0' WHERE ID=" & Session("UID") & ";"
	
	objConn.Execute(sqlverify)
	'Response.Write " " & objRec("Password") &"<br>"
	If objRec("Password") = Request.Form("OldPassword") Then
		objRec.close
		objRec.Open "UserInfo", objConn
		objConn.Execute(sql)
		Response.Redirect "main.asp"
	Else
		Response.Write "Your old password does not match the password you entered, please go back and try again."
		Response.Write "<BR><input type='button' onclick='javascript:history.go(-1)' value='<<Back'>"
		
	End If

	objRec.close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing	
	
%>
