<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"-->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim ThisPage
	Dim Level_ID

	Dim Page
	Page	= Request("Page")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<body>

<%
	Dim TrBgColor
	
	Level_ID = Request("Level_ID")
	CheckNum Level_ID
	Sql = "Select [Level_Name], [Level_Order] From [Monolith_Users_Level] Where [Level_ID]=" & Level_ID
	Rs.Open Sql, Conn, 1, 3
	If Rs.Bof or Rs.Eof then
	  Response.write "����ȼ������ڣ�"
		Response.end
	end if	
%>
<form name="form1" method="post" action="Level_Del_Submit.asp">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;��Ա���� &gt;&gt; ��Ա�ȼ�ɾ��</font></b></td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999" align="right">
			<font color="#FFFFFF">�������ƣ�</font></td>
			<td bgcolor="#FFFFFF"> &nbsp;<%=Rs(0)%> </td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">
			<font color="#FFFFFF">˳&nbsp;&nbsp;&nbsp; ��</font></td>
			<td bgcolor="#FFFFFF"> &nbsp;<%=Rs(1)%> </td>
		</tr>
		<tr>
			<td bgcolor="#999999" width="100" height="20" align="right">��</td>
			<td bgcolor="#FFFFFF">
			<input type="hidden" value="<%=Level_ID%>" name="Level_ID">
			<input type="submit" name="Submit" value=" ɾ �� (Alt + Y) " class="InputBox" accesskey="Y"><input type="hidden" name="Page" value="<%=Page%>">
			<input type="reset" value=" �� �� (Alt + N) " name="B2" class="InputBox" accesskey="N" onclick="window.open('Level.asp?Page=<%=Page%>','_self');">
			</td>
		</tr>
	</table>
</form>
<table><tr><td height="398"></td></tr></table>
<%
	Call CloseRs
	Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->