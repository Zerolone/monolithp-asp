<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.inc" -->
<!-- #include file="std.asp" -->
<%
	OnlyLogin
	Alive "ϵͳ��Ϣ"
%>
<HTML>
<head>
<meta http-equiv="pragma" content="no-cache" >
<link rel="stylesheet" type="text/css" href="main.css">
</head>
<%	dim rsSyslogin,rsBook
	dim todayLogNo,toWeekLogNo,toMonthLogno,TotalLogNo
	dim todayBook,toWeekBook,toMonthBook,TotalBook
	set rsSysLogin=server.CreateObject("ADODB.RecordSet")
	set rsBook=server.CreateObject("ADODB.RecordSet")
	
	rsSyslogin.Cachesize=1
	rsSyslogin.Open "select * from syslogin where DateDiff('d',loginDate, now)<1",myconn,1
	todayLogNo=rsSyslogin.RecordCount
	rsSyslogin.Close
	rsSyslogin.Open "select * from syslogin where DateDiff('ww',loginDate, now)<1",myconn,1
	toWeekLogNo=rsSyslogin.RecordCount
	rsSyslogin.Close
	rsSyslogin.Open "select * from syslogin where DateDiff('m',loginDate, now)<1",myconn,1
	toMonthLogNo=rsSyslogin.RecordCount
	rsSyslogin.Close
	rsSyslogin.Open "select * from syslogin",myconn,1
	totalLogNo=rsSyslogin.RecordCount
	rsSyslogin.Close
	
	rsBook.Open "select * from book where DateDiff('d',wdate,now)<1",myconn,1
	todayBook=rsBook.RecordCount
	rsBook.Close
	rsBook.Open "select * from book where DateDiff('ww',wdate,now)<1",myconn,1
	toWeekBook=rsBook.RecordCount
	rsBook.Close
	rsBook.Open "select * from book where DateDiff('m',wdate,now)<1",myconn,1
	toMonthBook=rsBook.RecordCount
	rsBook.Close
	rsBook.Open "select * from book",myconn,1
	totalBook=rsBook.RecordCount
	rsBook.Close

%>
<BODY bgcolor="#ccddff">
<hr width=80%>
<table width=90% border=0 align=center  cellpadding="2" cellspacing="1"   class=statInfo>
<tr>
	<td width=33% align=right >������վ�˴Σ�</td>
	<td width=17% align=left  ><%=todayLogNo%></td>
	<td width=33% align=right >������������</td>
	<td width=17% align=left  ><%=todayBook%></td>
</tr>
<tr>
	<td align=right>������վ�˴Σ�</td>
	<td><%=toWeekLogNo%></td>
	<td align=right>������������</td>
	<td><%=toWeekBook%></td>
</tr>
<tr>
	<td align=right>������վ�˴Σ�</td>
	<td><%=toMonthLogNo%></td>
	<td align=right>������������</td>
	<td><%=toMonthBook%></td>
</tr>
<tr>
	<td align=right>��&nbsp;&nbsp;��վ�˴Σ�</td>
	<td><%=totalLogNo%></td>
	<td align=right>��&nbsp;&nbsp;��������</td>
	<td><%=totalBook%></td>
</tr>
</table>
<hr width=80%>
<p>
</BODY>
</HTML>
<%	myconn.close
	set rsSysLogin=nothing
	set rsBook=nothing
	set myconn=nothing
%>
