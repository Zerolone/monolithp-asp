<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<%
	Dim Tree_ID, Tree_Parent_ID

	Tree_ID						= Request("Tree_ID")
	Tree_Parent_ID		= Request("Tree_Parent_ID")
	if Request("Del") = "True" then

		Sql	= "Delete From [Monolith_Tree] "
		Sql = Sql & " Where [Tree_ID]=" & Tree_ID
		Conn.execute(Sql)
		
		Sql	= "Select [Tree_ID] From [Monolith_Tree] Where [Tree_Parent_ID]=" & Tree_Parent_ID
		Set	Rs	= Conn.execute(Sql)
		if Rs.Bof Or Rs.Eof then
			Sql	= "Update [Monolith_Tree] Set "
			Sql = Sql & "[Tree_HasChild]=false "
			Sql = Sql & " Where [Tree_ID]=" & Tree_Parent_ID
			Conn.execute(Sql)
		end if
		Rs.Close
		Response.write "<script language=""javascript"">"
		Response.write "  alert(""ɾ���ɹ���"");"
		Response.write "  window.open('Default.asp','_self');"
		Response.write "</script>"
	end if
	

	'-----------------0-------------1----------------2------------3--------------4----------5---------------6-----------------
	Sql = "Select [Tree_ID], [Tree_Parent_ID], [Tree_Level], [Tree_Title], [Tree_Url], [Tree_Target], [Tree_HasChild] From [Monolith_Tree] Where [Tree_ID]= " & Tree_ID
	Rs.Open Sql, Conn , 1, 1
	if Not (Rs.Bof Or Rs.Eof) then
		if Rs(6) then
			Response.write "<script language=""javascript"">"
			Response.write "  alert(""�����ӽڵ㣬����ɾ��\n\n����ɾ���¼��ڵ㣡"");"
			Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
			Response.write "</script>"
		end if
%>
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<form method="POST" action="Del.asp?Del=True&Tree_ID=<%=Tree_ID%>&Tree_Parent_ID=<%=Rs(1)%>" name="FrmDelMenu">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;���˵����� &gt;&gt; �˵�ɾ��</font></b></td>
		</tr>			<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ID��</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<font color="#FF0000"><b><%=Rs(0)%></b></font></td>
		</tr>
		<tr <%if Rs(6) then Response.write "style="" display:none """  %>>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">��&nbsp; ��&nbsp; ��&nbsp; ID��</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(1)%></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� �� ����ȼ���</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(2)%></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� ��&nbsp;&nbsp; ��&nbsp; �⣺</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(3)%></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� ��&nbsp;&nbsp; ��&nbsp; �ӣ�</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(4)%></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� �� ���ӿ�ܣ�</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(5)%></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�Ƿ�����Ӳ˵���</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(6)%></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF"><input type="hidden" value="Same_Level" name="Tree_Level_Code">
			<input type="submit" value=" ɾ �� (Alt + Y) " name="B1" class="InputBox" accesskey="Y">
			<input type="button" value=" ȡ �� (Alt + N) " name="B2" onclick="window.open('Default.asp', '_self');" class="InputBox"></td>
		</tr>
	</table>
</form>
<%
	end if
	Call CloseRs
	Call CloseConn
%>
<table><tr><td height="300"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->