<%
	Dim objConn
	Dim objRec
	Dim sql
	Dim sql2

	
	Set objConn = Server.CreateObject ("ADODB.Connection")
	Set objRec = Server.CreateObject ("ADODB.Recordset")
	
	objConn.Open "DSN=HelpDesk2"
	objRec.Open "GROUP_MEMBERS", objConn
	
	sql = "UPDATE GROUP_MEMBERS SET  HW_SW_ID= '" & Request.Form("Category") & "', GROUP_NAME= '" & Request.Form("Group") & "', PACKAGE_NAME='" & Request.Form("Package") & "' WHERE PACK_ID=" & Request("Num") & ";"
	
	objConn.Execute(sql)
	
	objRec.close
	objConn.Close
	Set objRec = Nothing
	Set objConn = Nothing

	Response.Redirect "Index.asp"
%>
