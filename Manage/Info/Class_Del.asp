<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<%
	Dim Info_Class_ID, Info_Class_Parent_ID

	Info_Class_ID			= Request("Info_Class_ID")
	Info_Class_Parent_ID		= Request("Info_Class_Parent_ID")
	if Request("Del") = "True" then

		Sql	= "Delete From [Monolith_Info_Class] "
		Sql = Sql & " Where [Info_Class_ID]=" & Info_Class_ID
'		Response.write Sql
		Conn.execute(Sql)
		
		Sql	= "Select [Info_Class_ID] From [Monolith_Info_Class] Where [Info_Class_Parent_ID]=" & Info_Class_Parent_ID
'		Response.write Sql
		Set	Rs	= Conn.execute(Sql)
		if Rs.Bof Or Rs.Eof then
			Sql	= "Update [Monolith_Info_Class] Set "
			Sql = Sql & "[Info_Class_HasChild]=false "
			Sql = Sql & " Where [Info_Class_ID]=" & Info_Class_Parent_ID
			Conn.execute(Sql)
		end if
		Rs.Close
		Response.Redirect "Class_Setting.asp"
	end if
	

	'------------------------0-----------------------1----------------2-------------------3-------------------4------------------------5
	Sql = "Select [Info_Class_ID], [Info_Class_Parent_ID], [Info_Class_Level], [Info_Class_Title], [Info_Class_HasChild] , [Info_Users_Level] From [Monolith_Info_Class] Where [Info_Class_ID]= " & Info_Class_ID
	Rs.Open Sql, Conn , 1, 1
	if Not (Rs.Bof Or Rs.Eof) then
		if Rs(4) then
			Response.write "<script language=""javascript"">"
			Response.write "  alert(""存在子节点，不能删除\n\n请先删除下级节点！"");"
			Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
			Response.write "</script>"
		end if
%>
<form method="POST" action="Class_Del.asp?Del=True&Info_Class_ID=<%=Info_Class_ID%>&Info_Class_Parent_ID=<%=Rs(1)%>" name="FrmDelMenu">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;信息栏目管理 &gt;&gt; 栏目删除</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">栏&nbsp; 目&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ID：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<font color="#FF0000"><b><%=Rs(0)%></b></font></td>
		</tr>
		<tr <%if Rs(4) then Response.write "style="" display:none """  %>>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">父&nbsp; 栏&nbsp; 目&nbsp; ID：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(1)%></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">栏&nbsp; 目排序等级：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(2)%></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">栏&nbsp; 目&nbsp; 标&nbsp; 题：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(3)%></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">是否存在子栏目：</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(4)%></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">栏&nbsp; 目浏览权限：</font></td>
			<td bgcolor="#F6F6F6"><%
			Dim Rs1, i
			Set Rs1		= Server.CreateObject("ADODB.Recordset")
			'-------------------0-----------1
			Sql = "Select [Level_ID], [Level_Name] From [Monolith_Users_Level] Order By [Level_Order] Asc"
			Rs1.Open Sql, Conn ,1, 1
			Do While Not (Rs1.Bof Or Rs1.Eof)

				Dim Info_Users_Level_Arr
				Info_Users_Level_Arr	= Split(Rs(5), ",")
				For i = LBound(Info_Users_Level_Arr) to UBound(Info_Users_Level_Arr)
					if Cint(Trim(Info_Users_Level_Arr(i))) = Rs1(0) then Response.write Rs1(1) & "&nbsp;"
				Next

				Rs1.MoveNext
			Loop
			Rs1.Close
			Set Rs1=Nothing
		%>　</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF"><input type="hidden" value="Same_Level" name="Info_Class_Level_Code">
			<input type="submit" name="Submit0" value=" 删 除 (Alt + Y) " class="InputBox" accesskey="Y">
			<input type="button" value=" 取 消 (Alt + N)  " name="B2" class="InputBox" onclick="window.open('Class_Setting.asp','_self');" accesskey="N">
		</tr>
	</table>
</form>
<%
	end if
	Call CloseRs
	Call CloseConn
%>
<table>
	<tr>
		<td height="310"></td>
	</tr>
</table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->