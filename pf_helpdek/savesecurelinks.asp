<%
strLinkName = Request.Form("LinkName")
strLinkAddress = Request.Form("LinkAddress")

Set objConn = Server.CreateObject ("ADODB.Connection")
Set objRec = Server.CreateObject ("ADODB.Recordset")

objConn.Open "Helpdesk2"

objRec.Open "SecureLinks", objConn

sql = "INSERT INTO SecureLinks(LinkName, LinkAddress) VALUES ('"& strLinkName &"','"& strLinkAddress &"');"

objConn.Execute(sql)

Response.Redirect Request.ServerVariables("HTTP_Referer")
%>