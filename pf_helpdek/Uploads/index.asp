<%
If Request("Num") = "" Then
	strNum = "000"
Else
	strNum = Request("Num")
End If
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
<link rel="stylesheet" href="/helpdesk.css" type="text/css">
</HEAD>
<BODY>
<FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="uploadfile.asp?UserID=<%=strNum%>&RefPage=<%=Request("RefPage")%>">
	<TABLE BORDER=0>
	<tr><td><b>Select a file to upload:</b><br><INPUT TYPE=FILE SIZE=50 NAME="FILE1"></td></tr>
	<tr><td><INPUT TYPE=hidden NAME="saveto" value="disk"></td></tr>
	<tr><td align="center"><INPUT TYPE=SUBMIT VALUE="Upload File"></td></tr>
	
	</TABLE>
</FORM>
</BODY>
</HTML>
