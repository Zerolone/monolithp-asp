<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"--><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.css" rel="stylesheet" type="text/css">
<%
	Dim Template_ID, Template_Parent_ID

	Template_ID			= Request("Template_ID")
	Template_Parent_ID		= Request("Template_Parent_ID")
	if Request("Del") = "True" then

		Sql	= "Delete From [Monolith_Template] "
		Sql = Sql & " Where [Template_ID]=" & Template_ID
'		Response.write Sql
		Conn.execute(Sql)
		
		Sql	= "Select [Template_ID] From [Monolith_Template] Where [Template_Parent_ID]=" & Template_Parent_ID
'		Response.write Sql
		Set	Rs	= Conn.execute(Sql)
		if Rs.Bof Or Rs.Eof then
			Sql	= "Update [Monolith_Template] Set "
			Sql = Sql & "[Template_HasChild]=false "
			Sql = Sql & " Where [Template_ID]=" & Template_Parent_ID
			Conn.execute(Sql)
		end if
		Rs.Close
		Response.Redirect"Template_Setting.asp"
	end if
	

	'------------------------0-----------------------1----------------2-------------------3-------------------4
	Sql = "Select [Template_ID], [Template_Parent_ID], [Template_Level], [Template_Title], [Template_HasChild] From [Monolith_Template] Where [Template_ID]= " & Template_ID
	Rs.Open Sql, Conn , 1, 1
	if Not (Rs.Bof Or Rs.Eof) then
		if Rs(4) then
			Response.write "<script language=""javascript"">"
			Response.write "  alert(""�����ӽڵ㣬����ɾ��\n\n����ɾ���¼��ڵ㣡"");"
			Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
			Response.write "</script>"
		end if
%>
<form method="POST" action="Template_Del.asp?Del=True&Template_ID=<%=Template_ID%>&Template_Parent_ID=<%=Rs(1)%>" name="FrmDelMenu">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;ģ����� &gt;&gt; ��Ŀɾ��</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">��&nbsp;&nbsp;&nbsp; ��&nbsp;&nbsp;&nbsp; ID��</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<font color="#FF0000"><b><%=Rs(0)%></b></font></td>
		</tr>
		<tr <%if Rs(4) then Response.write "style="" display:none """  %>>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">��&nbsp;&nbsp; ��&nbsp; �� ID��</font></td>
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
			<font color="#FFFFFF">�Ƿ�����Ӳ˵���</font></td>
			<td bgcolor="#FFFFFF">&nbsp;<%=Rs(4)%></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF"><input type="hidden" value="Same_Level" name="Template_Level_Code">
			<input type="submit" value=" ȷ �� ɾ �� " name="B1" class="InputBox">
			<input type="button" value=" ȡ �� �� �� " name="B2" onclick="window.open('Template_Setting.asp','_self');" class="InputBox"></td>
		</tr>
	</table>
</form>
<%
	end if
	Call CloseRs
	Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->