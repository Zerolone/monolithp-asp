<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<%
	Dim News_Class_ID, News_Class_Parent_ID

	News_Class_ID			= Request("News_Class_ID")
	News_Class_Parent_ID		= Request("News_Class_Parent_ID")
	if Request("Del") = "True" then

		Sql	= "Delete From [Monolith_News_Class] "
		Sql = Sql & " Where [News_Class_ID]=" & News_Class_ID
'		Response.write Sql
		Conn.execute(Sql)
		
		Sql	= "Select [News_Class_ID] From [Monolith_News_Class] Where [News_Class_Parent_ID]=" & News_Class_Parent_ID
'		Response.write Sql
		Set	Rs	= Conn.execute(Sql)
		if Rs.Bof Or Rs.Eof then
			Sql	= "Update [Monolith_News_Class] Set "
			Sql = Sql & "[News_Class_HasChild]=false "
			Sql = Sql & " Where [News_Class_ID]=" & News_Class_Parent_ID
			Conn.execute(Sql)
		end if
		Rs.Close
		Response.write "<script language=""javascript"">"
		Response.write "  alert(""删除成功！"");"
		Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
		Response.write "</script>"
	end if
	

	'------------------------0-----------------------1----------------2-------------------3-------------------4
	Sql = "Select [News_Class_ID], [News_Class_Parent_ID], [News_Class_Level], [News_Class_Title], [News_Class_HasChild] From [Monolith_News_Class] Where [News_Class_ID]= " & News_Class_ID
	Rs.Open Sql, Conn , 1, 1
	if Not (Rs.Bof Or Rs.Eof) then
		if Rs(4) then
			Response.write "<script language=""javascript"">"
			Response.write "  alert(""存在子节点，不能删除\n\n请先删除下级节点！"");"
			Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
			Response.write "</script>"
		end if
%>
<form method="POST" action="Class_Del.asp?Del=True&News_Class_ID=<%=News_Class_ID%>&News_Class_Parent_ID=<%=Rs(1)%>" name="FrmDelMenu">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;新闻栏目管理 &gt;&gt; 栏目删除</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">菜单ID：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<font color="#FF0000"><b><%=Rs(0)%></b></font></td>
		</tr>
		<tr <%if Rs(4) then Response.write "style="" display:none """  %>>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">父菜单ID：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(1)%></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">菜单排序等级：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(2)%></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">菜单标题：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(3)%></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">是否存在子菜单：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(4)%></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF"><input type="hidden" value="Same_Level" name="News_Class_Level_Code">
			<input type="submit" value=" 确 认 删 除 " name="B1" class="InputBox">
			<input type="button" value=" 取 消  " name="B2" class="InputBox" onclick="window.open('Class_Setting.asp','_self');"></td>
		</tr>
	</table>
</form>
<%
	end if
	Call CloseRs
	Call CloseConn
%>