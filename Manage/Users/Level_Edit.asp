<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Manage/Check.asp"--><%
	Dim ThisPage
	Dim Level_ID

	Level_ID = Request("Level_ID")
	CheckNum Level_ID

	Dim Page
	Page	= Request("Page")
	
  ThisPage					= Request.ServerVariables("PATH_INFO")
	if Request("Code") = "Edit" then
		'-------------------0-------------1
		Sql = "Select [Level_Name], [Level_Order] From [Monolith_Users_Level] Where [Level_ID]=" & Level_ID
		Rs.Open Sql, Conn, 1, 3
		if Not(Rs.Bof Or Rs.Eof) then
			Rs(0)		= Request("Level_Name")
			Rs(1)		= Request("Level_Order")
			Rs.Update
		end if
		Rs.Close

		Response.write("<script language='javascript'>alert('修改成功！')</script>")
	end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<body>

<%
	Sql = "Select [Level_Name], [Level_Order] From [Monolith_Users_Level] Where [Level_ID]=" & Level_ID
	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof or Rs.Eof then
	  Response.write "这个广告不存在！"
		Response.end
	end if	
%>
<form name="form1" method="post" action="<%=ThisPage%>?Code=Edit">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="8"><b><font color="#FFFFFF">&nbsp;会员设置 &gt;&gt; 会员等级修改</font></b></td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">标题名称：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Level_Name" type="text" id="Level_Name" value="<%=Rs(0)%>" size="40" class="InputBox">
			</td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">
			<font color="#FFFFFF">顺&nbsp;&nbsp;&nbsp; 序：</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="Level_Order" class="InputBox" style="width=250">
			<%
				Dim i
				For i =1 to 20
			%>
			<option value="<%=i%>" <% if Rs(1) = i then Response.write "selected" %>><%=i%></option>
			<%Next%>
			</select></td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">　</td>
			<td colspan="2" bgcolor="#FFFFFF">
			<input type="hidden" name="Page" value="<%=Page%>">
			<input type="submit" name="Submit0" value=" 修 改 (Alt + Y) " class="InputBox" accesskey="Y">
			<input type="hidden" name="Level_ID" value="<%=Level_ID%>">
			<input type="button" value=" 取 消 (Alt + N) " name="B2" onclick="window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);" class="InputBox" accesskey="N">
			</td>
		</tr>
	</table>
</form>
<%
	Call CloseRs
	Call CloseConn
%><table><tr><td height="398"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->