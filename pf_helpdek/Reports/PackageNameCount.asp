<!--#include file="../_head.asp"-->
<td>


<%
Param = Request.QueryString("Param")
Data = Request.QueryString("Data")
%>
<%
If IsObject(Session("helpdesk_conn")) Then
    Set conn = Session("helpdesk_conn")
Else
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.open "helpdesk2","",""
    Set Session("helpdesk_conn") = conn
End If
%>
<%
    sql = "SELECT CALLS.Package_Name, Count(CALLS.TICKET_NUMBER) AS CountOfTICKET_NUMBER  FROM CALLS  "
    If cstr(Param) <> "" And cstr(Data) <> "" Then
        sql = sql & " WHERE [" & cstr(Param) & "] = " & cstr(Data)
    End If
    sql = sql & " GROUP BY CALLS.Package_Name, CALLS.CLOSED  HAVING (((CALLS.CLOSED)='NO'))  ORDER BY Count(CALLS.TICKET_NUMBER) DESC    "
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open sql, conn, 3, 3

    strTotal=0
%>
<TABLE BORDER=1 BGCOLOR=#ffffff CELLSPACING=0 align=left><B></B>

<THEAD>
<TR>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 >Person Working Call</TH>
<TH BGCOLOR=#c0c0c0 BORDERCOLOR=#000000 >Number of Open Tickets</TH>

</TR>
</THEAD>
<TBODY>
<%
On Error Resume Next
rs.MoveFirst
do while Not rs.eof

strPackage_Name= Server.HTMLEncode(rs.Fields("Package_Name").Value)
Set objConn=Server.CreateObject("ADODB.Connection")
objConn.Open "DSN=HelpDesk2"
Set objRs=objConn.Execute("Select * from Group_Members where Package_Name= '"&strPackage_Name&"'")

 %>
<TR VALIGN=TOP>

<TD BORDERCOLOR=#c0c0c0 >
<a href="workerticketlist.asp?worker=<%=strWorker%>">
<%
If Server.HTMLEncode(rs.Fields("Package_Name").Value) = "" Then
	Response.Write "Unassigned"
Else
	Response.Write objRs("Package_Name")
End If
%>
</a>
<BR></FONT></TD>
<TD BORDERCOLOR=#c0c0c0  ALIGN=RIGHT><%=Server.HTMLEncode(rs.Fields("CountOfTICKET_NUMBER").Value)%><BR></FONT></TD>

</TR>
<%
strTotal = strTotal + rs.Fields("CountOfTICKET_NUMBER").Value
rs.MoveNext
loop%>
<TR VALIGN=TOP>
<TD BORDERCOLOR=#c0c0c0 >
Total
<BR></FONT></TD>
<TD BORDERCOLOR=#c0c0c0  ALIGN=RIGHT><%=strTotal%><BR></FONT></TD>

</TR>
</TBODY>
<TFOOT></TFOOT>
</TABLE>

</BODY>
</HTML>