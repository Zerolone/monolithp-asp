<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<!-- #include file="std.asp" -->
<%
	OnlySysOp
	Alive "用户管理"
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
		alert("用户ID不能为空");
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
<center>请先查询用户，然后再进行管理，<br>或者直接在用户列表中对用户进行管理<br>
<form name=frmaLogin action="queryUser.asp">
用户ID：<INPUT TYPE="text" width=10 NAME="userid" class=font1>  <INPUT TYPE="button" class=font1 value="查询" onclick=check()>
</form>
<A HREF="userList.asp">用户管理列表</a>
<p>
</center>
</body>
</html>
<%	myconn.Close
	set myconn=nothing
%>
