<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<!-- #include file="std.asp" -->
<!-- #include file="string.asp" -->
<%
	Alive "��ѯ�û�"
	if request("userid")="" then
%>
<html >
<head>
<meta http-equiv="pragma" content="no-cache" >
<link rel="stylesheet" type="text/css" href="main.css">
<SCRIPT LANGUAGE="JavaScript">
<!--//
function check()
{
	if (document.frmaLogin.userid.value.length<1) {
		alert("�û�ID����Ϊ��");
		document.frmaLogin.userid.focus();
	}
	else
	{	
		document.frmaLogin.submit();
	}
}
//-->
</SCRIPT>
</head>
<body class=clblue>
<br><p>
<center>
<form name=frmaLogin action="queryUser.asp">
�û�ID��<INPUT TYPE="text" width=10 NAME="userid" class=font1>  <INPUT TYPE="button" class=font1 value="��ѯ" onclick=check()>
</form>
<p>
</center>
</body>
</html>
<%	else
		dim rsUserInfo,strSQL,iUser,i
		dim rsBook,iShow
		set rsUserInfo=server.CreateObject("ADODB.RecordSet")
		set rsBook=server.CreateObject("ADODB.RecordSet")
		rsUserInfo.Open "select * from UserInfo where userid='"+request("userid")+"'",myconn,1
		if rsUserInfo.RecordCount <1 then
			session("strURL")="queryUser.asp"
			session("strMsg")="�Ҳ����û�ID��"+request("userid")
			response.redirect "alert.asp"
		else%>
			<html >
			<head>
			<meta http-equiv="pragma" content="no-cache" >
			<link rel="stylesheet" type="text/css" href="main.css">
			<body class=clblue>
			<center><div align=center>
					<b><big>��ѯ���</big></b><br><br>
			<table width=90% class=queryUser border=1 cellpadding="0" cellspacing="0">
				<% if isSysOP(session("userid")) then%>
				<tr><td width=27%>&nbsp;&nbsp;����</td><td>
					&nbsp;&nbsp;
					<a href="adduser.asp?method=modify&userid=<%=rsUserInfo("userid")%>">�޸��û�����</a>
					&nbsp;&nbsp;&nbsp;&nbsp;
					<a href="adduser.asp?method=delete&userid=<%=rsUserInfo("userid")%>">ɾ���û�</a>
				</td></tr>
				<% end if%>
				<tr><td width=27%>&nbsp;&nbsp;�û�ID</td><td><%=rsUserInfo("UserID")%></td></tr>
				<tr><td>&nbsp;&nbsp;�س�</td><td><%=rsUserInfo("nickname")%></td></tr>
			    <tr><td>&nbsp;&nbsp;Email</td><td><%=rsUserInfo("Email")%></td></tr>
				<tr><td>&nbsp;&nbsp;��վ����</td><td><%=rsUserInfo("iLogin")%>��</td></tr>
				<tr><td>&nbsp;&nbsp;����������</td><td><%=rsUserInfo("iBook")%>ƪ����</td></tr>
				<tr><td>&nbsp;&nbsp;����ֵ</td><td><%=rsUserInfo("iPerience")%>��</td></tr>
				<tr><td>&nbsp;&nbsp;�����վʱ��</td><td><%=rsUserInfo("LastLogin")%></td></tr>
				<tr><td>&nbsp;&nbsp;ǩ����</td><td><%=showbody(rsUserInfo("sign")&" ")%>&nbsp;</td></tr>

			</table><br>
			<%	rsBook.CursorLocation = 3
				rsBook.CacheSize = 10
                rsBook.PageSize = 10
				rsBook.Open "select * from Book where userid='"+rsUserInfo("userid")+"' order by wdate desc",myconn,1
				iShow=10
				if rsBook.RecordCount>0 then %>
					<table width=90% border=1  cellpadding="0" cellspacing="0" class=queryUser >
					<tr><td align=center colspan=2>��������ʮƪ����<br></td></tr>
					<%  if iShow>rsBook.RecordCount then iShow=rsBook.RecordCount
						for i=1 to iShow %>
						<tr><td width=5%>&nbsp;<%=i%></td><td><a class=listtopic1 href="showdoc.asp?method=new&bid=<%=rsBook("bid")%>&pageNo=1&id=<%=rsBook("id")%>&reply=0">
					    <%=rsBook("title")%></a></td></tr>
					<%  rsBook.MoveNext
						next %>
					</table><br>
			<%  end if
				rsBook.Close
			
				rsBook.CursorLocation = 3
				rsBook.CacheSize = 10
                rsBook.PageSize = 10
				rsBook.Open "select * from SysLogin where userid='"+rsUserInfo("userid")+"' order by LoginDate desc",myconn,1
				iShow=10
				if rsBook.RecordCount>0 then %>
					<table width=90% border=1  cellpadding="0" cellspacing="0" class=queryUser >
					<tr><td align=center colspan=3>���ʮ����վIP��ʱ��<br></td></tr>
					<%  if iShow>rsBook.RecordCount then iShow=rsBook.RecordCount
						for i=1 to iShow %>
						<tr><td width=5%>&nbsp;<%=i%></td>
						    <td width=25%>&nbsp;&nbsp;<%=rsBook("ip")%></td>
							<td >&nbsp;&nbsp;<%=rsBook("LoginDate")%></td>
						</tr>
					<%  rsBook.MoveNext
						next %>
					</table>
			<%  end if%>
			</div></center>
			</body>
			</html>
<%		end if
	end if
%>
