<link rel="stylesheet" href="helpdesk.css" type="text/css">
<script>
<!-- Script to return selected user info to Add.asp.
function remote2(url){
	opener.location=url
	window.close()
}
function SearchFocus(){
	UserLookup.UserQuery.focus();
}
//-->
</script>

<%
Response.Write "<HTML><HEAD><TITLE>Select User</TITLE></HEAD><BODY link=black vlink=black alink=black onLoad='SearchFocus()'>"
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=Helpdesk2"

Logon=Request.Form("Logon_Name") + Request.Form("Domain")

'This section pulls the user information to be used to display a hello message.
'"Sort By" Section
'************************
If Request.Form("UserQuery") = "" Then
	strUserQuery = ""
	If Request("sortby")="LName" or Request("sortby")=""  then
		strSearch = "Last Name"
		'Set objRs=objConn.Execute("Select * from UserInfo ORDER by LastName, FirstName")
		strFilter = "Off"
	End if
	If Request("sortby")="FName" then
		strSearch = "First Name"
		'Set objRs=objConn.Execute("Select * from UserInfo ORDER by FirstName, LastName")
		strFilter = "Off"
	End if
	If Request("sortby")="Department"  then
		strSearch = "Department"
		'Set objRs=objConn.Execute("Select * from UserInfo ORDER by Department, LastName")
		strFilter = "Off"
	End if 
	If Request("sortby")="Country"  then
		strSearch = "Country"
		'Set objRs=objConn.Execute("Select * from UserInfo ORDER by Country, LastName")
		strFilter = "Off"
	End if
Else
'***************
	strUserQuery = Request.Form("UserQuery")
	If Request("sortby")="LName" or Request("sortby")=""  then
		strSearch = "Last Name"
		Set objRs=objConn.Execute("Select * from UserInfo WHERE (LastName LIKE '" & strUserQuery & "%') ORDER by LastName, FirstName;")
		strFilter = "On"
	End if
	If Request("sortby")="FName" then
		strSearch = "First Name"
		Set objRs=objConn.Execute("Select * from UserInfo WHERE (FirstName LIKE '" & strUserQuery & "%') ORDER by FirstName, LastName")
		strFilter = "On"
	End if
	If Request("sortby")="Department"  then
		strSearch = "Department"
		Set objRs=objConn.Execute("Select * from UserInfo WHERE (Department LIKE '" & strUserQuery & "%') ORDER by Department, LastName")
		strFilter = "On"
	End if 
	If Request("sortby")="Country"  then
		strSearch = "Country"
		Set objRs=objConn.Execute("Select * from UserInfo WHERE (Country LIKE '" & strUserQuery & "%') ORDER by Country, LastName")
		strFilter = "On"
	End if
End If
'************************
Response.Write "<table width='100%' border=0><tr><td align='left'>"
If strFilter = "Off" Then
	Response.Write "<FORM NAME='UserLookup' method='Post' action='userwindow.asp?sortby=" & Request("sortby") & "'><input type=text name='UserQuery' size='30'><BR>"
	Response.Write "<INPUT TYPE=SUBMIT Value='Search by " & strSearch & "'></FORM>"
Else
	Response.Write "<Form method='post' action='userwindow.asp?sortby=" & Request("sortby") & "'><input type=submit value='Turn Off Filter'></Form>"
End If
Response.Write "</td><td align=right valign='top'>"
Response.Write "<FORM><input type=Submit value='Close' onClick='javascript:window.close()'></Form>"
Response.Write "</td></tr></table>"

Response.Write "<HR>"
Response.Write "<table border=1 cellpadding=2 cellspacing=0 BORDERCOLORLIGHT=#CCCCCC BORDERCOLORDARK=#CCCCCC BORDERCOLOR=#CCCCCC width='100%'>"
Response.Write "<tr bgcolor=#CCCCCC NOWRAP  BORDERCOLORDARK=White BORDERCOLORLIGHT=Black>"

'************************
'Determine if table is sorted by Last Name and make the column header a link to sort by if it is not.
If strSearch = "Last Name" Then
	Response.Write "<td NOWRAP>Last Name</td>"
Else
	Response.Write "<td NOWRAP><a href=userwindow.asp?sortby=LName>Last Name</a></td>"
End If

'************************
'Determine if table is sorted by First Name and make the column header a link to sort by if it is not.
If strSearch = "First Name" Then
	Response.Write "<TD NOWRAP>First Name</TD>"
Else
	Response.Write "<td NOWRAP><a href=userwindow.asp?sortby=FName>First Name</a></TD>"
End If

'************************
'Determine if table is sorted by Department and make the column header a link to sort by if it is not.
If strSearch = "Department" Then
	Response.Write "<TD NOWRAP>Department</TD>"
Else
	Response.Write "<TD><a href=userwindow.asp?sortby=Department>Department</TD>"
End If

'************************
'Determine if table is sorted by Country and make the column header a link to sort by if it is not.
If strSearch = "Country" Then
	Response.Write "<TD NOWRAP>Country</TD>"
Else
	Response.Write "<TD><a href=userwindow.asp?sortby=Country>Country</TD>"
End If


Response.Write "</tr>"		
If Request.Form("UserQuery") = "" Then
Else
	While Not objRs.EOF
		Response.Write "<tr>"
		Response.Write "<td><a href=javascript:remote2('Add.asp?Num=" & objRs("ID") & "')>" & objRs("LastName") &"</a></td>"
		Response.Write "<td NOWRAP>" & objRs("FirstName") & "</td>"
		Response.Write "<td NOWRAP>" & objRs("Department") & "</td>"
		Response.Write "<td NOWRAP>" & objRs("Country") & "</td>"
		Response.Write "</tr>"	
		objRs.MoveNext
	Wend
End IF
	Response.Write "</table>"
If Request.Form("UserQuery") = "" Then
Else
objRs.close
Set objRs= Nothing
objConn.Close
set objConn=nothing
End IF

	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"


%>
<P>
<Form method='get' action='adduser.asp'><input type=submit value='Add User'></Form>
</body>
</html>
