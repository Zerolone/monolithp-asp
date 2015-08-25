<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
</head>

<body>

<%
	Dim Advertisement_ID
	Advertisement_ID = Request("Advertisement_ID")
	CheckNum Advertisement_ID
	
	'---------------------0----------------------1---------------------2---------------------------3-----------------4------------------------5----------------------6-------------------------7-----------------------8----------------------9-----------------------10-------------------11---------------------12
	Sql = "SELECT [Advertisement_ID], [Advertisement_Title], [Advertisement_Hits], [Advertisement_Width], [Advertisement_Height], [Advertisement_Class_ID], [Advertisement_Active], [Advertisement_FileName], [Advertisement_Url], [Advertisement_Content], [Advertisement_Color], [Advertisement_No], [Advertisement_Order] FROM [Monolith_Advertisement] Where [Advertisement_ID]=" & Advertisement_ID
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
		<font color="#FFFFFF">U&nbsp; r&nbsp; l：</font></td>
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
