<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Manage/Check.asp"-->
<!--#include virtual="/Inc/MD5.asp"-->

<%
	Dim ThisPage
	Dim Admin_ID
	
	Dim Page
	Page		= Request("Page")

  ThisPage					= Request.ServerVariables("PATH_INFO")
	if Request("Code") = "Add" then
		'-------------------0-----------1-------------2
		Sql = "Select [Admin_ID], [Admin_Name], [Admin_Password] From [Monolith_Admin]"
		Rs.Open Sql, Conn, 1, 3
		Rs.AddNew
		Rs(1)		= Request("Admin_Name")
		Rs(2)		= MD5(Request("Admin_Password"), 32)
		Rs.Update
		Admin_ID	= Rs(0)
		Call CloseRs
		Call CloseConn
		
		Response.Redirect "Edit.asp?Admin_ID=" & Admin_ID
		Response.end
	end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<body>
<form name="form1" method="post" action="<%=ThisPage%>?Code=Add">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;管理员管理 &gt;&gt; 添加管理员</font></b></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">用户名：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Admin_Name" type="text" id="Links_Hits" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">密&nbsp; 码：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Admin_Password" type="text" id="Links_Hits" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20">　</td>
			<td bgcolor="#FFFFFF">
			<input type="submit" name="Submit" value=" 添 加 管 理 员 " class="InputBox">
			<input type="button" name="Submit2" value=" 返 回 " onclick="window.open('Default.asp', '_self');" class="InputBox">
			</td>
		</tr>
	</table>
</form>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->

</body>

</html>