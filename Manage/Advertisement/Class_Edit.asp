<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"--><!-- #include virtual ="/Manage/Check.asp"--><%
	Dim ThisPage
	Dim Advertisement_Class_ID, Cmd

	Dim Page
	Page	= Request("Page")
	
  ThisPage					= Request.ServerVariables("PATH_INFO")
	if Request("Code") = "Edit" then
		Advertisement_Class_ID = Request("Advertisement_Class_ID")
		CheckNum Advertisement_Class_ID
		'-----------------------0------------------------------1
		Sql = "Select [Advertisement_Class_Name], [Advertisement_Class_Order] From [Monolith_Advertisement_Class] Where [Advertisement_Class_ID]=" & Advertisement_Class_ID
		Rs.Open Sql, Conn, 1, 3
		Rs(0)		= Request("Advertisement_Class_Name")
		Rs(1)		= Request("Advertisement_Class_Order")
		Rs.Update
		Call CloseRs
		Call CloseConn

		Response.Redirect "Class.asp?Page=" & Page
	end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<body>

<%
	Dim TrBgColor
	
	Advertisement_Class_ID = Request("Advertisement_Class_ID")
	CheckNum Advertisement_Class_ID
	Sql = "Select [Advertisement_Class_Name], [Advertisement_Class_Order] From [Monolith_Advertisement_Class] Where [Advertisement_Class_ID]=" & Advertisement_Class_ID
	Rs.Open Sql, Conn, 1, 3
	If Rs.Bof or Rs.Eof then
	  Response.write "这个广告不存在！"
		Response.end
	end if	
%>
<form name="form1" method="post" action="<%=ThisPage%>?Code=Edit">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="8"><b><font color="#FFFFFF">&nbsp;广告类型管理 &gt;&gt; 广告类型修改</font></b></td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">标题名称：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Advertisement_Class_Name" type="text" id="Advertisement_Class_Name" value="<%=Trim(Rs(0))%>" size="40" class="InputBox">
			<input type="hidden" name="Advertisement_Class_Name_Old" value="<%=Trim(Rs(1))%>">
			</td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">
			<font color="#FFFFFF">顺&nbsp;&nbsp;&nbsp; 序：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Advertisement_Class_Order" type="text" id="Advertisement_Class_Order" value="<%=Rs(1)%>" size="40" class="InputBox">
			</td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">　</td>
			<td colspan="2" bgcolor="#FFFFFF">
			<input type="hidden" name="Page" value="<%=Page%>">
			<input type="submit" name="Submit0" value=" 修 改 类 型 (Alt + A) " class="InputBox" accesskey="A">
			<input type="hidden" name="Advertisement_Class_ID" value="<%=Advertisement_Class_ID%>">
			<input type="reset" value=" 返 回 (Alt + N) " name="B2" class="InputBox" accesskey="N" onclick="window.open('Class.asp?Page=<%=Page%>','_self');">
			</td>
		</tr>
	</table>
</form>
<%
	Call CloseRs
	Call CloseConn
%><table><tr><td height="398"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->