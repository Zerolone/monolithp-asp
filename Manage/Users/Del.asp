<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Users_ID
	Users_ID = Request("Users_ID")
	CheckNum Users_ID

	Dim Page
	Page	= Request("Page")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<body>
<%
	Dim TrBgColor
	
	Users_ID = Request("Users_ID")
	CheckNum Users_ID
	'-------------------0-------------1--------------2
	Sql = "Select [Users_Name], [Users_Group], [Users_LastLogin] From [Monolith_Users] Where [Users_ID]=" & Users_ID
	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof or Rs.Eof then
	  Response.write "这个广告不存在！"
		Response.end
	end if	

%>
<form name="form1" method="post" action="Del_Submit.asp">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2">
			<b><font color="#FFFFFF">&nbsp;会员管理 &gt;&gt; 删除会员</font></b></td>
		</tr>  <tr>
    <td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">会员名称：</font></td>
    <td bgcolor="#FFFFFF">
      &nbsp;<%=Rs(0)%>
    </td>
  </tr>
  <tr>
    <td height="20" bgcolor="#999999" width="100" align="right">
			<font color="#FFFFFF">会员等级：</font></td>
    <td bgcolor="#FFFFFF">
      &nbsp;<%=Rs(1)%>
    </td>
    </tr>
  <tr>
    <td bgcolor="#999999" width="100" height="20" align="right">
			<font color="#FFFFFF">最后登陆：</font></td>
    <td bgcolor="#FFFFFF">
      &nbsp;<%=Rs(2)%>
    </td>
    </tr>
  <tr>
    <td bgcolor="#999999" width="100" height="20" align="right">　</td>
    <td bgcolor="#FFFFFF">
			<input type="submit" name="Submit0" value=" 删 除 (Alt + Y) " class="InputBox" accesskey="Y">
			<input type="reset" value=" 返 回 (Alt + N) " name="B2" class="InputBox" accesskey="N" onclick="window.open('Default.asp?Page=<%=Page%>','_self');">
			<input type="hidden" name="Users_ID" value="<%=Users_ID%>">
    <input type="hidden" name="Page" value="<%=Page%>">
  </tr>
</table>
</form>
<%
		Call CloseRs
		Call CloseConn
%><table><tr><td height="380"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->