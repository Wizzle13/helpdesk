
<%
host = Request.Form("host")
If host = "" Then host = "www.yahoo.com"
%>

<form method=post>

<center>
<table width="90%">
<tr><td>Host Name/Address: <td><input type=text name=host value="<% =host %>" size=50>
<tr><td>Timeout (seconds): <td><input type=text name=timeout value="10" size=5>
<tr><td>Resolve Names:     <td><input type=checkbox name=resolve> (This could take a long time - a larger timeout may be needed.)
<tr><td><td><input type=submit value="   Trace Now!   ">
</table>
</center>

</form>

<%
If Request("REQUEST_METHOD") = "POST" Then
  
  Set traceroute = Server.CreateObject("IPWorksASP.TraceRoute")

  traceroute.Timeout = Request.Form("timeout")
  
  If Request.Form("resolve") = "on" Then traceroute.ResolveNames = True
  
  On Error Resume Next
  
  traceroute.TraceTo host
  
  If Err.Number <> 0 Then
    Response.Write "<hr><font color=red><b>Error: " & Err.Description
    Response.Write " (the table below contains the partial route obtained so far)</font>"
  End If
  
  On Error GoTo 0
%>

<hr>
<center>
<table width="90%">
  <tr>
    <th>Hop</th>
    <th>Hop Address</th>
    <th>Hop Host Name</th>
    <th>Hop Time (ms)</th>
  </tr>
  
<% For i=1 To traceroute.HopCount %>

  <tr>
    <td><% =i %></td>
    <td><% =traceroute.HopHostAddress(i) %></td>
    <td><% =traceroute.HopHostName(i) %></td>
    <td><% =traceroute.HopTime(i) %></td>
  </tr>
  
<% Next %>

</table>
</center>

<%
End If
%>
