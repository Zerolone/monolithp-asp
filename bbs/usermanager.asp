<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<!-- #include file="std.asp" -->
<%
	OnlySysOp
	Alive "�û�����"
%>
<html >
<head>
<meta http-equiv="pragma" content="no-cache" >
<link rel="stylesheet" type="text/css" href="main.css">
</head>
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

<body class=clblue>
<%		dim rsSysInfo
		dim showSysop,strSQL,method
%>
<br><p>
<center>���Ȳ�ѯ�û���Ȼ���ٽ��й���<br>����ֱ�����û��б��ж��û����й���<br>
<form name=frmaLogin action="queryUser.asp">
�û�ID��<INPUT TYPE="text" width=10 NAME="userid" class=font1>  <INPUT TYPE="button" class=font1 value="��ѯ" onclick=check()>
</form>
<A HREF="userList.asp">�û������б�</a>
<p>
</center>
</body>
</html>
<%	myconn.Close
	set myconn=nothing
%>
