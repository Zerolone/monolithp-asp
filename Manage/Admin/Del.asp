<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim ThisPage
	Dim Admin_ID
	Admin_ID	= Request("Admin_ID")
	CheckNum Admin_ID
	
	Dim Page
	Page		= Request("Page")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<body>

<%
	'-------------------0-------------1
	Sql = "Select [Admin_Name], [Admin_Password] From [Monolith_Admin] Where [Admin_ID]=" & Admin_ID
	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof or Rs.Eof then
	  Response.write "�������Ա�����ڣ�"
		Response.end
	end if	

%>
<form name="form1" method="post" action="Del_Submit.asp">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;���ӹ��� &gt;&gt; ɾ������</font></b></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">�û�����</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Admin_Name" type="text" value="<%=Rs(0)%>" size="40" class="InputBox" disabled></td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" width="100" height="20">
			<font color="#FFFFFF">��&nbsp; �룺</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Admin_Password" type="text" value="<%=Rs(1)%>" size="40" class="InputBox" disabled></td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20">��</td>
			<td bgcolor="#FFFFFF">
			<input type="hidden" name="Page" value="<%=Page%>">
			<input type="hidden" name="Admin_ID" value="<%=Admin_ID%>">
			<input type="submit" name="Submit" value=" ɾ �� �� �� Ա " class="InputBox">
			<input type="button" name="Submit2" value=" �� �� " onclick="window.open('Default.asp', '_self');" class="InputBox">
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