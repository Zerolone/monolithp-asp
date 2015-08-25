<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
</head>

<body>

<%
	Dim Users_ID
	Users_ID = Request("Users_ID")
	CheckNum Users_ID
	
	'-------------------0-------------1--------------2------------------3
	Sql = "Select [Users_Name], [Users_Group], [Users_LastLogin], [Users_Active] From [Monolith_Users] Where [Users_ID]=" & Users_ID

	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof or Rs.Eof then
	  Response.write "这个广告不存在！"
		Response.end
	end if	
%> <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="2">
		<b><font color="#FFFFFF">&nbsp;广告管理 &gt;&gt; 广告察看</font></b></td>
	</tr>
	<tr>
		<td bgcolor="#999999" align="right" width="100" height="20">
		<font color="#FFFFFF">标&nbsp;&nbsp;&nbsp; 题：</font></td>
		<td bgcolor="#FFFFFF">
&nbsp;<%=Rs(1)%></td>
	</tr>
	<tr>
		<td bgcolor="#999999" align="right" width="100" height="20">
		<font color="#FFFFFF">点&nbsp;&nbsp;&nbsp; 击：</font></td>
		<td bgcolor="#FFFFFF">
&nbsp;<%=Rs(2)%></td>
	</tr>
	<tr>
		<td bgcolor="#999999" align="right" width="100" height="20">
		<font color="#FFFFFF">长&nbsp;&nbsp;&nbsp; 度：</font></td>
		<td bgcolor="#FFFFFF">
&nbsp;<%=Rs(3)%></td>
	</tr>
	<tr>
		<td bgcolor="#999999" align="right" width="100" height="20">
		<font color="#FFFFFF">宽&nbsp;&nbsp;&nbsp; 度：</font></td>
		<td bgcolor="#FFFFFF">
&nbsp;<%=Rs(4)%></td>
	</tr>
	<tr>
		<td bgcolor="#999999" align="right" width="100" height="20">
		<font color="#FFFFFF">所&nbsp;&nbsp;&nbsp; 属：</font></td>
		<td bgcolor="#FFFFFF">
&nbsp;<%=Rs(5)%></td>
	</tr>
	<tr>
		<td bgcolor="#999999" align="right" width="100" height="20">
		<font color="#FFFFFF">使&nbsp;&nbsp;&nbsp; 用：</font></td>
		<td bgcolor="#FFFFFF">
&nbsp;<%=Rs(6)%> 　</td>
	</tr>
	<tr>
		<td bgcolor="#999999" align="right" width="100" height="20">
		<font color="#FFFFFF">文 件 名：</font></td>
		<td bgcolor="#FFFFFF">
&nbsp;<%=Rs(7)%> 　</td>
	</tr>
	<tr>
		<td bgcolor="#999999" align="right" width="100" height="20">
		<font color="#FFFFFF">U&nbsp; R&nbsp; L ：</font></td>
		<td bgcolor="#FFFFFF">
&nbsp;<%=Rs(8)%> 　</td>
	</tr>
	<tr>
		<td bgcolor="#999999" align="right" width="100" height="20">
		<font color="#FFFFFF">背景颜色：　　</font></td>
		<td bgcolor="#FFFFFF">
&nbsp;<%=Rs(10)%> 　</td>
	</tr>
	<tr>
		<td bgcolor="#999999" align="right" width="100" height="20">
		<font color="#FFFFFF">随 机 ID：　</font></td>
		<td bgcolor="#FFFFFF">
&nbsp;<%=Rs(11)%> 　</td>
	</tr>
	<tr>
		<td bgcolor="#999999" align="right" width="100" height="20">
		<font color="#FFFFFF">顺&nbsp;&nbsp;&nbsp; 序： </font></td>
		<td bgcolor="#FFFFFF">
&nbsp;<%=Rs(12)%> 　</td>
	</tr>
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="2">
		<b><font color="#FFFFFF">&nbsp;广告管理 &gt;&gt; 广告效果察看</font></b></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td colspan="2"><%=Rs(9)%></td>
	</tr>
</table>
<table><tr><td height="<%=224-Rs(4)%>"></td></tr></table>
<%
	Call CloseRs
	Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->

</body>

</html>
