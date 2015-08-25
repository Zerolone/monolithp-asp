<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<!-- #include file="std.asp" -->
<%
	OnlyLogin
	Alive "系统信息"
%>
<HTML>
<head>
<meta http-equiv="pragma" content="no-cache" >
<link rel="stylesheet" type="text/css" href="main.css">
</head>
<%	dim rsSyslogin,i
	dim todayLogNo,toWeekLogNo,toMonthLogno,TotalLogNo
	set rsSysLogin=server.CreateObject("ADODB.RecordSet")

	
	rsSyslogin.Cachesize=1
	rsSyslogin.Open "select * from syslogin where DateDiff('d',loginDate, now)=0",myconn,1
	todayLogNo=rsSyslogin.RecordCount
	rsSyslogin.Close
	rsSyslogin.Open "select * from syslogin where DateDiff('ww',loginDate, now)=0",myconn,1
	toWeekLogNo=rsSyslogin.RecordCount
	rsSyslogin.Close
	rsSyslogin.Open "select * from syslogin where DateDiff('m',loginDate, now)=0",myconn,1
	toMonthLogNo=rsSyslogin.RecordCount
	rsSyslogin.Close
	rsSyslogin.Open "select * from syslogin",myconn,1
	totalLogNo=rsSyslogin.RecordCount
	rsSyslogin.Close
	

%>
<BODY class=clblue>
<hr width=80%>
<table width=90% border=0 align=center  cellpadding="2" cellspacing="1"   class=statInfo>
<tr>
	<td width=33% align=right >当天上站人次：</td>
	<td width=17% align=left  ><%=todayLogNo%></td>
</tr>
<tr>
	<td align=right>本周上站人次：</td>
	<td><%=toWeekLogNo%></td>
</tr>
<tr>
	<td align=right>本月上站人次：</td>
	<td><%=toMonthLogNo%></td>
</tr>
<tr>
	<td align=right>总&nbsp;&nbsp;上站人次：</td>
	<td><%=totalLogNo%></td>
</tr>
</table>
<div align=center>
<TABLE bgColor=#000000 border=0 cellPadding=0 cellSpacing=1 width=500>
<TR>
  <TD align=middle bgColor=white>
    <FONT class=p9><B>最近十天上站人数</B></FONT>
  </TD>
</TR>
<TR><TD>
  <TABLE bgColor=#ffffee cellPadding=0 cellSpacing=0 height=80 width=100% border=0>
<%	dim ndate,tBook(11),tmax,twidth
	tmax=1
	for i=1 to 10
		rsSysLogin.Open "select * from syslogin where DateDiff('d',loginDate, now)="+cstr(i-1),myconn,1
		tBook(i)=rsSysLogin.RecordCount
		if tmax<tbook(i) then tmax=tbook(i)
		rsSysLogin.Close
	next
	for i=1 to 10
		twidth=int((tBook(i)/tmax)*70)
		ndate=right(" "&cstr(datepart("d",DateAdd("d",-1*(i-1), now))),2)
%>
	  <TR><TD align=left height=10>&nbsp;<%=ndate%>日：
		<IMG align=absMiddle height=10 src="image/bar.jpg" width=<%=twidth%>%>
		<%=tbook(i)%>
	  </TD></TR>
<%  next %>
  </TABLE>
</TD></TR>
</TABLE>
</div>
<hr width=80%>
</BODY>
</HTML>
<%	myconn.close
	set rsSysLogin=nothing
	set myconn=nothing
%>