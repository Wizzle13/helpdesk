<form method=post action=color.asp>
Blue: <input name=blue type=text><BR>
Green: <input name=green type=text><BR>
Red: <input name=red type=text><BR>
<input type=submit value=Calculate>
<%
Dim BGRvalue

BGRvalue = (Request.Form("blue") * 65536) + (Request.Form("green") * 256) + Request.Form("red")

Response.Write "<input type=text value=" & BGRvalue & "></form>"

%>