<!--#include file="_head.asp"-->
<td valign=top>
<script language="Javascript">
	function checknumber(){
		var x=document.TickNum.Num.value
		var anum=/(^\d+$)|(^\d+\.\d+$)/
		if (anum.test(x))
		testresult=true
	else{
		alert("That's not a number. Please Enter a Number.")
		testresult=false
	}
	return (testresult)
}

	function checkban(){
		if (document.layers||document.all)
			return checknumber()
		else
			return true
}
</script>
<% 
Session("PageName") = "Search"
	Response.Write "<table border=1>"
	

	Response.Write "<tr><td valign=top>Ticket Number:</td>"
	Response.Write "<td><form method=post action=Viewticket.asp onSubmit='return checkban()' name='TickNum'>"
	Response.Write "<input type='text' name='Num' size='15'>"
	Response.Write "<input type=Submit value=Search></form></td></tr>"

'-----> Keyword Search <------'	
	Response.Write "<tr><td valign=top>Keyword or Phrase:</td>"
	Response.Write "<td><form method=post action=userkeywordsearch.asp>"
	Response.Write "<input type='text' size='15' name='Keyword'><input type=Submit value=Search>"
	Response.Write "<BR><Select name=KeywordSearch>"
	Response.Write "<Option selected value=Exact>The exact phrase entered<option value=All>All of the words entered"
	Response.Write "</form></td></tr>"
'******************************
 %>
</table>
</td>
</tr>
</table>

</body>
</html>
