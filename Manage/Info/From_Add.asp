<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Inc/MD5.asp"-->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim ThisPage
	Dim From_ID
	From_ID = Request("From_ID")
	CheckNum From_ID	
		
	if Request("Code") = "Add" then
		Dim From_Name, From_Order
		
		From_Name		= Request("From_Name")
		From_Order	= Request("From_Order")

		'--------------------0--------------1
		Sql = "Select [From_Name], [From_Order] From [Monolith_Info_From]"
		Rs.Open Sql, Conn, 1, 3
		Rs.AddNew
		Rs(0)		= From_Name
		Rs(1)		= From_Order
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
<form name="form1" method="post" action="From_Add.asp?Code=Add">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;��Ϣ����ϵͳ���� &gt;&gt; ��Ϣ</font></b><font color="#FFFFFF"><b>��Դ���� 
		&gt;&gt; ��Ϣ��Դ���</b></font></td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">��Ϣ���ߣ�</font></td>
			<td bgcolor="#FFFFFF">
			<input name="From_Name" type="text" size="40" class="InputBox" style="width=250"></td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">˳&nbsp;&nbsp;&nbsp; ��</font></td>
			<td bgcolor="#FFFFFF">
			<select name="From_Order" size="1" class="InputBox" style="width=250">
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