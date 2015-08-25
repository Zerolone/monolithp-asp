<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<%
	Dim ThisPage
  ThisPage					= Request.ServerVariables("PATH_INFO")
	if Request("Code") = "Add" then
		'-----------------------0------------------------------1
		Sql = "Select [Advertisement_Class_Name], [Advertisement_Class_Order] From [Monolith_Advertisement_Class]"
		Rs.Open Sql, Conn, 1, 3
		Rs.AddNew
		Rs(0)		= Request("Advertisement_Class_Name")
		Rs(1)		= Request("Advertisement_Class_Order")
		Rs.Update
		Call CloseRs
		Call CloseConn

		Response.Redirect("Class.asp")
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
		<td colspan="8"><b><font color="#FFFFFF">&nbsp;广告类型管理 &gt;&gt; 广告类型添加</font></b></td>
	</tr>  
  <tr>
    <td bgcolor="#999999" height="20" align="right" width="100"><font color="#FFFFFF">标题名称：</font></td>
    <td bgcolor="#FFFFFF">
      <input name="Advertisement_Class_Name" type="text" id="Advertisement_Class_Name" size="40" class="InputBox">
    </td>
  </tr>
  <tr>
    <td bgcolor="#999999" height="20" align="right" width="100"><font color="#FFFFFF">顺&nbsp;&nbsp;&nbsp; 序：</font></td>
    <td bgcolor="#FFFFFF">
      <input name="Advertisement_Class_Order" type="text" id="Advertisement_Class_Order" size="40" class="InputBox">
    </td>
  </tr>
  <tr>
    <td bgcolor="#999999" height="20" align="right" width="100">　</td>
    <td bgcolor="#FFFFFF">
			<input type="submit" name="Submit0" value=" 添 加 类 型 (Alt + A) " class="InputBox" accesskey="A"></td>
  </tr>
</table>
</form>
<table><tr><td height="398"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->