<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.inc" -->
<% if session("userid")<>"" then%>
<%	dim rsUserInfo
	dim strSQL
	set rsUserInfo=server.createobject("ADODB.RECORDSET")
%>
<html><head><title><%=Application("Title")%></title>
<link rel="stylesheet" type="text/css" href="main.css">
</head>
<body class=clblue>
<table class=listTopic width=70%>
<tr><td width=20%>用户ID</td><td width=30%>IP</td><td width=50%>上站日期</td></tr>
<%	
	dim i,iNO
	strSQL="select * from syslogin where loginOk=-1 order by logindate desc"
	rsUserInfo.Open strSQL,myconn,1
	iNo=rsUserInfo.RecordCount
	for i=1 to iNo
%>
	<tr><td><%=rsUserInfo("userid")%></td><td><%=rsUserInfo("ip")%></td><td><%=rsUserInfo("loginDate")%></td></tr>
<%	rsUserInfo.MoveNext
	next %>
</table>
<BODY>
</BODY></HTML>
<% 
else
	response.redirect "quit.htm"
end if%>