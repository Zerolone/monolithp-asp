<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Inc/MD5.asp"-->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim ThisPage
	Dim Author_ID
	Author_ID = Request("Author_ID")
	CheckNum Author_ID	
		
	Dim Page

	Page	= Request("Page")

  ThisPage					= Request.ServerVariables("PATH_INFO")
	if Request("Code") = "Edit" then
		Dim Author_Name, Author_Order
		
		Author_Name		= Request("Author_Name")
		Author_Order	= Request("Author_Order")

		'--------------------0--------------1
		Sql = "Select [Author_Name], [Author_Order] From [Monolith_Info_Author] Where [Author_ID]=" & Author_ID
		Rs.Open Sql, Conn, 1, 3
		Rs(0)		= Author_Name
		Rs(1)		= Author_Order
		Rs.Update
		Rs.Close

		Response.write("<script language='javascript'>alert('�޸ĳɹ���')</script>")

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
		
	'--------------------0--------------1
	Sql = "Select [Author_Name], [Author_Order] From [Monolith_Info_Author] Where [Author_ID]=" & Author_ID
	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof or Rs.Eof then
	  Response.write "������߲����ڣ�"
		Response.end
	end if
%>
<form name="form1" method="post" action="<%=ThisPage%>?Code=Edit">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;��Ϣ����ϵͳ���� &gt;&gt; ��Ϣ����</font></b><font color="#FFFFFF"><b>���� 
		&gt;&gt; ��Ϣ�����޸�</b></font></td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">��Ϣ���ߣ�</font></td>
			<td bgcolor="#FFFFFF">
			<input name="Author_Name" type="text" size="40" class="InputBox" value="<%=Rs(0)%>" style="width=250"></td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">˳&nbsp;&nbsp;&nbsp; ��</font></td>
			<td bgcolor="#FFFFFF">
			<select name="Author_Order" size="1" class="InputBox" style="width=250">
			<%For i=1 to 100%>
			<option value="<%=i%>" <% if Rs(1) = i then Response.Write "selected" %>><%=i%></option>
			<%Next%>
			</select> </td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">��</td>
			<td bgcolor="#FFFFFF">
			<input type="submit" name="Submit0" value=" �� �� (Alt + Y) " class="InputBox" accesskey="Y">
			<input type="button" value=" ȡ �� (Alt + N) " name="B2" onclick="window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);" class="InputBox">
			<input type="hidden" name="Author_ID" value="<%=Author_ID%>">
			<input type="hidden" name="Page" value="<%=Page%>"></td>
		</tr>
	</table>
</form>
<%
		Call CloseRs
		Call CloseConn
%><table><tr><td height="330"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->