<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Inc/MD5.asp"-->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim ThisPage
	Dim KeyWord_ID
	KeyWord_ID = Request("KeyWord_ID")
	CheckNum KeyWord_ID	
		
	if Request("Code") = "Add" then
		Dim KeyWord_Name, KeyWord_Order
		
		KeyWord_Name		= Request("KeyWord_Name")
		KeyWord_Order	= Request("KeyWord_Order")

		'--------------------0--------------1
		Sql = "Select [KeyWord_Name], [KeyWord_Order] From [Monolith_News_KeyWord]"
		Rs.Open Sql, Conn, 1, 3
		Rs.AddNew
		Rs(0)		= KeyWord_Name
		Rs(1)		= KeyWord_Order
		Rs.Update
		Call CloseRs
		Call CloseConn

		Response.write "<script language=""javascript"">"
		Response.write "  alert(""��ӳɹ�, ��ˢ��ҳ��쿴����ӵ����ݡ�"");"
		Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
		Response.write "</script>"
	end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<body>

<%
	Dim i
%>
<form name="form1" method="post" action="KeyWord_Add.asp?Code=Add">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;����ϵͳ���� &gt;&gt; ���Źؼ���</font></b><font color="#FFFFFF"><b>���� 
		&gt;&gt; ���Źؼ������</b></font></td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">���Źؼ��֣�</font></td>
			<td bgcolor="#FFFFFF">
			<input name="KeyWord_Name" type="text" size="40" class="InputBox" style="width=250"></td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">˳&nbsp;&nbsp;&nbsp; ��</font></td>
			<td bgcolor="#FFFFFF">
			<select name="KeyWord_Order" size="1" class="InputBox" style="width=250">
			<%For i=1 to 100%>
			<option value="<%=i%>"><%=i%></option>
			<%Next%>
			</select> </td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">��</td>
			<td bgcolor="#FFFFFF">
			<input type="submit" name="Submit0" value=" �� �� (Alt + Y) " class="InputBox" accesskey="Y">
			<input type="button" value=" ȡ �� (Alt + N) " name="B2" onclick="window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);" class="InputBox"></td>
		</tr>
	</table>
</form>
<table><tr><td height="330"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->