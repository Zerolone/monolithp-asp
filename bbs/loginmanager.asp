<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<!-- #include file="std.asp" -->
<%
  OnlySysOp
  Alive "系统管理"
     dim rsUserInfo,strSQL,StrOrder,strAct
	 set rsUserInfo=server.createobject("ADODB.RECORDSET")

	 strAct=request("Action")
	 if strAct="clearAll" then
		strSQL="DELETE * from syslogin"
		rsUserInfo.Open strSQL,myconn,1
	 end if
	 if strAct="clearError" then 
		strSQL="DELETE * from syslogin where loginOk=0 "
		rsUserInfo.Open strSQL,myconn,1
	 end if
	 

	 strOrder=request("SortOrder")
	 if strOrder="" or ucase(strorder)="DATE" then 
		strOrder="logindate desc"
	 else if ucase(strOrder)="IP" then
			strOrder="ip asc"
	      else
			strOrder="userid asc,logindate desc"
	      end if 
	 end if 
%>
<html><head>
<% if strAct="clearError" or strAct="clearAll" then %>
<script language="javascript">
	alert("日志已清除！");
</script>
<% end if%>
<title><%=Application("Title")%></title>
<link rel="stylesheet" type="text/css" href="main.css">
</head>
<body class=clblue>
<center><a href=loginmanager.asp?action=clearAll>[清除上站日志]</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=loginmanager.asp?action=clearError>[清除错误上站日志]</a></center><br>
<table class=listTopic width=60% align=center >
<tr><td width=20%><a href=loginmanager.asp?sortorder=userid >用户ID</a></td><td width=30%><a href=loginmanager.asp?sortorder=IP >IP</a></td><td width=50%><a href=loginmanager.asp?sortorder=DATE >上站日期</a></td></tr>
<%	
	dim i,iNO
	strSQL="select * from syslogin where loginOk=-1 order by "+strOrder
	rsUserInfo.Open strSQL,myconn,1
	iNo=rsUserInfo.RecordCount
	for i=1 to iNo
%>
	<tr><td><%=rsUserInfo("userid")%></td><td><%=rsUserInfo("ip")%></td><td><%=rsUserInfo("loginDate")%></td></tr>
<%	rsUserInfo.MoveNext
	next
	rsUserInfo.close%>
</table>
<br>
<hr>
<center>---错误登录日志---</center><br>
<table class=listTopic width=60% align=center>
<tr><td width=20%><a href=loginmanager.asp?sortorder=userid >用户ID</a></td><td width=30%><a href=loginmanager.asp?sortorder=IP >IP</a></td><td width=50%><a href=loginmanager.asp?sortorder=DATE >上站日期</a></td></tr>
<%	
	strSQL="select * from syslogin where loginOk=0 order by "+strOrder
	rsUserInfo.Open strSQL,myconn,1
	iNo=rsUserInfo.RecordCount
	for i=1 to iNo
%>
	<tr><td><%=rsUserInfo("userid")%></td><td><%=rsUserInfo("ip")%></td><td><%=rsUserInfo("loginDate")%></td></tr>
<%	rsUserInfo.MoveNext
	next
	rsUserInfo.Close
%>
</table>
<BODY>
</BODY></HTML>
<%
	myconn.Close
	set rsUserInfo=nothing
	set myconn=nothing
%>