<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim ThisPage
  ThisPage					= Request.ServerVariables("PATH_INFO")
	if Request("Code") = "Add" then
		'-------------------0-------------1
		Sql = "Select [Level_Name], [Level_Order] From [Monolith_Users_Level]"
		Rs.Open Sql, Conn, 1, 3
		Rs.AddNew
		Rs(0)		= Request("Level_Name")
		Rs(1)		= Request("Level_Order")
		Rs.Update
		Call CloseRs
		Call CloseConn

		Response.write "<script language=""javascript"">"
		Response.write "  alert(""��ӳɹ�, ��ˢ��ҳ��쿴����ӵ����ݡ�"");"
		Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
		Response.write "</script>"

		Response.end
	end if
%>
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<body>

<form name="form1" method="post" action="<%=ThisPage%>?Code=Add">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;��Ա���� &gt;&gt; ��Ա�ȼ����</font></b></td>
		</tr>
		<tr>
			<td bgcolor="#999999" height="20" align="right" width="100">
			<font color="#FFFFFF">��Ա�ȼ���</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Level_Name" type="text" size="40" class="InputBox">
			</td>
		</tr>
		<tr>
			<td bgcolor="#999999" height="20" align="right" width="100">
			<font color="#FFFFFF">��Ա˳��</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="Level_Order" class="InputBox" style="width=250">
			<%
				Dim i
				For i =1 to 20
			%>
			<option value="<%=i%>"><%=i%></option>
			<%Next%>
			</select></td>
		</tr>
		<tr>
			<td bgcolor="#999999" height="20" align="right" width="100">��</td>
			<td bgcolor="#FFFFFF">
			<input type="submit" name="Submit0" value=" �� �� (Alt + Y) " class="InputBox" accesskey="Y">
			<input type="button" value=" ȡ �� (Alt + N) " name="B2" onclick="window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);" class="InputBox" accesskey="N"></td>
		</tr>
	</table>
</form>
<table>
	<tr>
		<td height="398"></td>
	</tr>
</table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->

</body>

</html>