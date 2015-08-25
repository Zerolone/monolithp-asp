<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Manage/Check.asp"--><%
	Dim ThisPage
	Dim Links_ID
	Links_ID	= Request("Links_ID")
	CheckNum Links_ID
	
	Dim Page
	Page		= Request("Page")

  ThisPage					= Request.ServerVariables("PATH_INFO")
	if Request("Code") = "Edit" then
		'-------------------0--------------1------------2--------------3------------4-------------5
		Sql = "Select [Links_Title], [Links_Url], [Links_Order], [Links_Pic], [Links_Kind], [Links_Active] From [Monolith_Links] Where [Links_ID]=" & Links_ID
		Rs.Open Sql, Conn, 1, 3
		Rs(0)		= Request("Links_Title")
		Rs(1)		= Request("Links_Url")
		Rs(2)		= Request("Links_Order")
		Rs(3)		= Request("Links_Pic")
		Rs(4)		= Request("Links_Kind")
		Rs(5)		= Request("Links_Active")
		Rs.Update
		Call CloseRs
		Call CloseConn
		
		Response.Redirect "Default.asp?Page=" & Page
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

<%
	'-------------------0--------------1------------2--------------3------------4-------------5
	Sql = "Select [Links_Title], [Links_Url], [Links_Order], [Links_Pic], [Links_Kind], [Links_Active] From [Monolith_Links] Where [Links_ID]=" & Links_ID
	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof or Rs.Eof then
	  Response.write "这个链接不存在！"
		Response.end
	end if	

%>
<form name="form1" method="post" action="<%=ThisPage%>?Code=Edit">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;链接管理 &gt;&gt; 修改链接</font></b></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">名称：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Links_Title" type="text" value="<%=Rs(0)%>" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">地址：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Links_Url" type="text" value="<%=Rs(1)%>" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">顺序：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Links_Order" type="text" value="<%=Rs(2)%>" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">图片：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Links_Pic" type="text" id="Links_Hits0" value="<%=Rs(3)%>" size="40" class="InputBox"></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">类型：</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="Links_Kind" class="InputBox">
			<option value="1" <% if rs(4)=1 then response.write("selected") %>>文字链接
			</option>
			<option value="2" <% if rs(4)=2 then response.write("selected") %>>图片链接
			</option>
			</select></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">使用：</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Links_Active" type="radio" value="0" id="Links_Active_0" <% if not rs(5) then response.write "checked" %>><label for="Links_Active_0">不使用</label>
			<input name="Links_Active" type="radio" value="1" id="Links_Active_1" <% if rs(5) then response.write "checked" %>><label for="Links_Active_1">使用</label> 
			　</td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20">　</td>
			<td bgcolor="#FFFFFF">
			<input type="hidden" name="Page" value="<%=Page%>">
			<input type="hidden" name="Links_ID" value="<%=Links_ID%>">
			<input type="submit" name="Submit" value=" 修 改 链 接 " class="InputBox">
			<input type="button" name="Submit2" value=" 返 回 " onclick="window.open('Default.asp', '_self');" class="InputBox">
			</td>
		</tr>
	</table>
</form>
<%
		Call CloseRs
		Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->

</body>

</html>
