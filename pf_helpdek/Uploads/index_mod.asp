<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
<link rel="stylesheet" href="/helpdesk.css" type="text/css">
<script>
function status(){
	document.clear();
	document.writeln("<P>&nbsp</P><center><img src=/images/uploadstatus.gif></center>");
	form.submit();
}
</script>
</HEAD>
<BODY>
<FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="uploadfile_mod.asp?TicketNum=<%=Request("TicketNum")%>&RefPage=<%=Request("RefPage")%>">
	<TABLE BORDER=0>
	<tr><td><b>Select a file to upload:</b><br><INPUT TYPE=FILE SIZE=50 NAME="FILE1"></td></tr>
	<tr><td><INPUT TYPE=hidden NAME="saveto" value="disk"></td></tr>
	<tr><td align="center"><INPUT TYPE=SUBMIT VALUE="Upload File"></td></tr>
	
	</TABLE>
</FORM>
</BODY>
</HTML>
