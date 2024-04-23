<!--#include file="../_validsession.asp"-->
<html>
<head><title>Online Help Desk Chat</title></head>
<SCRIPT LANGUAGE="VBS">
</SCRIPT>

<body onLoad=Flux>

<OBJECT classid=clsid:D6526FE0-E651-11CF-99CB-00C04FD64497 codeBase=MSChatOCX.Cab#Version=4,71,730,0 codeType=application/x-oleobject height=375 id=Chat width=575 STANDBY = "Dowloading the Microsoft MSChat ActiveX Control">
	<SPAN STYLE="color:red">Online Help Desk chat failed to automatically install!<P><a href="/chat/chatinstall.bat">Click here to install the chat software manually</a>.<P></SPAN><SPAN STYLE="color:black">Please note the following:<BR>1. When it asks you if you want to Save or Run the file, click Run, then click OK.<BR>2. You will see 2 dialog boxes pop up during the install, click OK on both.<BR>3. You will have to close this window and re-open it to gain full access to Online Help Chat.</SPAN>
	<PARAM NAME="_ExtentX" VALUE="15875">
	<PARAM NAME="_ExtentY" VALUE="6615">
	<PARAM NAME="MaxMessageLength" VALUE="2000">
	<PARAM NAME="MaxHistoryLength" VALUE="32767">
	<PARAM NAME="Appearance" VALUE="3">
	<PARAM NAME="BorderStyle" VALUE="0">
	<PARAM NAME="UIOption" VALUE="2815">
	<PARAM NAME="BackColor" VALUE="255">
</OBJECT>

<SCRIPT LANGUAGE="VBS">
SUB Flux
	Dim Form
   	Set Form = Document.WEBChat

	If (Chat.State = 1) Then
		Chat.EnterRoom "mic://172.16.2.21/#InformationTechnology", "", Form.Alias.Value, "ANON", 9, 1
	Else 
		If (Chat.State = 2) Then
			Chat.CancelEntering
			Chat.ClearHistory
		Else 
			If (Chat.State = 3) Then
				Chat.ExitRoom
				Chat.ClearHistory
			End If
		End If
	End If
END SUB

SUB Chat_OnStateChanged(ByVal NewState)
	If (NewState = 1) Then
		'Document.WEBChat.FluxBtn.Value = "   Join the chat   "
		Chat.BackColor = 255
	Else
		If (NewState = 2) Then
			'Document.WEBChat.FluxBtn.Value = "Cancel Entering"
			Chat.BackColor = 16777215
		Else
			If (NewState = 3) Then
				'Document.WEBChat.FluxBtn.Value = "Leave the chat"
				Chat.BackColor = 12632256
			End If
		End If
	End If
END SUB
</SCRIPT>
<FORM NAME=WEBChat>
<INPUT NAME="Alias" Type=hidden VALUE="<%Response.Write Session("Email")%>">
</FORM>
</body>
</html>