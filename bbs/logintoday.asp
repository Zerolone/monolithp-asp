<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<!-- #include file="std.asp" -->
<%
  OnlySysOp
  Alive "系统管理"
  
     dim rsUserInfo,strSQL,i
	 set rsUserInfo=server.createobject("ADODB.RECORDSET")
	 rsUserInfo.Open "select userid,count(*)  as ilogin from syslogin where datediff('d',loginDate, now)=0 group by userid order by userid ",myconn,1
%>
<html><head>
<link rel="stylesheet" type="text/css" href="main.css">
</head>
<body class=clblue>
<div align=center ><center>
<br>
  <table width=70% class=InfoMan border=1 cellpadding="0" cellspacing="0">
  <tr><td colspan=2 align=center>当日上站人次</td></tr>
  <% for i=1 to rsUserInfo.RecordCount %>
  <tr><td width=50% >&nbsp;<A href="queryUser.asp?userid=<%=rsUserInfo("UserID")%>"><%=rsUserInfo("UserID")%></a></td>
      <td>&nbsp;&nbsp;<%=rsUserInfo("iLogin")%>
  </td></tr>
  <%	rsUserInfo.MoveNext
	 next
     rsUserInfo.Close
  %>
  </table>
  </center></div>
</BODY></HTML>
<%
	myconn.Close
	set rsUserInfo=nothing
	set myconn=nothing
%>