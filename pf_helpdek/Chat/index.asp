<!--#include file="../_validsession.asp"-->
<script>
function remote2(url){
	opener.location=url
	window.close()
}
</script>

<html>
<head><title>Online Help Desk Chat</title></head>
<SCRIPT LANGUAGE="VBS">
</SCRIPT>

<body>

Welcome to Online Help Desk Chat.  While in the online chat, please remember these notes:
<P>
<OL>
<LI>This is for Help Desk issues only.  Anyone using this for other reasons will be kicked out of the chat immediately.
<LI>You <U>must</U> have an open ticket before entering the chat room.  We will NOT help you if you do not have an open issue.  Please be prepared to give us the ticket number when asked.
</OL>
Do you have an open ticket with the Help Desk?<BR>
<FORM NAME=ValidTick action=chat.asp>
<INPUT NAME="Yes" Type=submit VALUE="Yes, I have an open ticket to discuss.">
<INPUT NAME="No" Type=button VALUE="No, I need to open a ticket." onClick="remote2('/open.asp')">
</FORM>
</body>
</html>