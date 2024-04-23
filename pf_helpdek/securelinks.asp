<!--#include file="_head.asp"-->
<body>
<%
Dim strWorkID
strWorkID = Session("WorkID")

Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")
Set objRs = Server.CreateObject ("ADODB.Recordset")

objConn.Open "Helpdesk2"

Set objRec=objConn.Execute("SELECT * FROM UserInfo WHERE WorkerID='" & strWorkID & "';")

Set objRs=objConn.Execute("Select * FROM Departments WHERE Department_Name='" & objRec("Department") & "';")

Response.Write "<td valign=top>"
Response.Write "<P>&nbsp;</P>"

If objRs("DepartmentID") = "ITSYS" Then

'*************************Put Secure Links here.**********************
Set objRs2=objConn.Execute("Select * FROM SecureLinks ORDER BY LinkName;")
Response.Write "<P>"
Response.Write "<UL>"

While Not objRs2.EOF
	Response.Write "<LI><a href='" & objRs2("LinkAddress") & "' target=new>" & objRs2("LinkName") & "</A> - (<a href=deletesecurelinks.asp?Num=" & objRs2("LinkID") & " alt='Delete' class=pw><font color=red>Delete</font></a>)"
	objRs2.MoveNext
Wend

Response.Write "</UL>"
Response.Write "<P>&nbsp;</P>"

Response.Write "<B>Add New Link</B>"
Response.Write "<form method=post action=savesecurelinks.asp>"
Response.Write "Link Name: <input type=textbox size=25 name=LinkName><BR>"
Response.Write "Link Address: <input type=file size=50 name=LinkAddress><BR>"
Response.write "<input type=submit value='Add Link'> * Please include 'http://' when posting Internet links."


'*************************End of secure Links.************************
Else
	Response.Write "<B><font color=red>You do not have authorization to view this page.  If you feel you have received this message in error, please contact the IT Help Desk USA.</font></B>"
End If

Response.Write "</td>"
Response.Write "</tr>"
Response.Write "</table>"
%>
</body>
</html>