<%
'Checks to see if the session is still valid.
If Session("IsValid") = "False" or Session("IsValid") = "" Then
	Session("URL") = Request.ServerVariables("SCRIPT_NAME") + "?" + Request.ServerVariables("QUERY_STRING")
	Response.Redirect "/logoff.asp"
Else
	Response.Redirect "/main.asp"
End If
%>
