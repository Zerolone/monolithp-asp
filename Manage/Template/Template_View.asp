<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript" src="../Js/All.js"></script>
<script language=javascript>
function GoToClass()
{
	var ClassValue, ClassText;
	ClassValue	= News_Class.value;
	ClassText		= News_Class.options[News_Class.selectedIndex].text;
	ClassText		= ClassText.replace(ClassValue, "");
	ClassText		= ClassText.replace("-", "");
	ClassText		= ClassText.substring(ClassValue.length, ClassText.length);
	ffClick("News/News_List.asp?News_Class=" + ClassValue, ClassText);
}

function GoToClass()
{
	var ClassValue, ClassText;
	ClassValue	= News_Class.value;
	ClassText		= News_Class.options[News_Class.selectedIndex].text;
	ClassText		= ClassText.replace(ClassValue, "");
	ClassText		= ClassText.replace("-", "");
	ClassText		= ClassText.substring(ClassValue.length, ClassText.length);
	ffClick("News/News_List.asp?News_Class=" + ClassValue, ClassText);
}
</script>
<br>　<table border="1" width="100%" id="table1">
	<tr>
		<td>
	<p align="center">
	<select size="26" name="News_Class" class="InputBox"  style="width=300" onchange="GoToClass();">
	<option value="">察看所有新闻</option>
	<%
		'---------------------0-----------------1------------------2
		Sql	= "Select [News_Class_ID], [News_Class_Level], [News_Class_Title] From [Monolith_News_Class] Order By [News_Class_Level]"
		Response.write Sql
		Rs.Open Sql, Conn, 1, 1
		Do While Not (Rs.Bof Or Rs.Eof)
		%>
		<option value="<%=Rs(1)%>"><%LoopNBSP(Len(Rs(1)))%><%=Rs(1)%>-<%=Rs(2)%></option>
		<%
			Rs.MoveNext
		Loop
		Call CloseRs
		Call CloseConn
	%>
	</select><br>
<input type="button" value=" 察 看 栏 目 新 闻 " name="B3" onclick="GoToClass();" class="InputBox" style="width=300">

</td>
		<td>
	<p align="center">
	<select size="26" name="News_Display" class="InputBox"  style="width=300" onchange="GoToDisplay();">
	<option value="1">首页调用新闻</option>
	</select><br><input type="button" value=" 察 看 板 块 新 闻 " name="B4" onclick="GoToDisplay();" class="InputBox" style="width=300"></td>
	</tr>
</table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->