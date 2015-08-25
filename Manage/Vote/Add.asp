<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"--><!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<body>

<%
	Dim ThisPage
  ThisPage					= Request.ServerVariables("PATH_INFO")
	if Request("Code") = "Add" then
		'------------------0-------------1---------------2---------------3---------------4---------------5---------------6---------------7---------------8
		Sql = "Select [Vote_Title], [Vote_Answer1], [Vote_Answer2], [Vote_Answer3], [Vote_Answer4], [Vote_Answer5], [Vote_Answer6], [Vote_Answer7], [Vote_Answer8],"
		'------------------9
		Sql	= Sql & " [Vote_Active] From [Monolith_Vote]"
		Rs.Open Sql, Conn, 1, 3
		Rs.AddNew
		Rs(0)		= Request("Vote_Title")
		Rs(1)		= Request("Vote_Answer1")
		Rs(2)		= Request("Vote_Answer2")
		Rs(3)		= Request("Vote_Answer3")
		Rs(4)		= Request("Vote_Answer4")
		Rs(5)		= Request("Vote_Answer5")
		Rs(6)		= Request("Vote_Answer6")
		Rs(7)		= Request("Vote_Answer7")
		Rs(8)		= Request("Vote_Answer8")
		Rs(9)		= Request("Vote_Active")
		Rs.Update
		Call CloseRs
		Call CloseConn
		
		Response.Redirect "Default.asp"
	end if
%>
<form name="form1" method="post" action="<%=ThisPage%>?Code=Add">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="3"><b><font color="#FFFFFF">&nbsp;调查管理 &gt;&gt; 添加调查项目</font></b></td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">标&nbsp;&nbsp;&nbsp; 题：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Vote_Title" type="text" id="Vote_Title" size="40" class="InputBox">
			</td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">问 题 一：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Vote_Answer1" type="text" id="Vote_Hits" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">问 题 二：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Vote_Answer2" type="text" id="Vote_Hits" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">问 题 三：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Vote_Answer3" type="text" id="Vote_Hits" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">问 题 四：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Vote_Answer4" type="text" id="Vote_Hits" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">问 题 五：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Vote_Answer5" type="text" id="Vote_Hits" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">问 题 六：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Vote_Answer6" type="text" id="Vote_Hits" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">问 题 七：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Vote_Answer7" type="text" id="Vote_Hits" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">问 题 八：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Vote_Answer8" type="text" id="Vote_Hits" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">
			<font color="#FFFFFF">使&nbsp;&nbsp;&nbsp; 用：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Vote_Active" type="radio" value="0" id="Vote_Active_0"><label for="Vote_Active_0">不使用</label>
			<input type="radio" name="Vote_Active" value="1" id="Vote_Active_1"><label for="Vote_Active_1">使用</label>
			</td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20">　</td>
			<td colspan="2" bgcolor="#FFFFFF">
			<input type="submit" name="Submit" value=" 添 加 调 查 项 目 " class="InputBox">
			</td>
		</tr>
	</table>
</form>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->

</body>

</html>