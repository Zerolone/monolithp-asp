<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<title>收到新消息</title>
<style type="text/css">
td{font-size:9pt;}
A:link {COLOR: #000000; TEXT-DECORATION: none; font-size: 9pt}
A:active {COLOR: #000000; TEXT-DECORATION: none; font-size: 9pt}
A:visited {COLOR: #000000; TEXT-DECORATION: none; font-size: 9pt}
.song9 {  font-family: 宋体; font-size: 9pt}
</style>
<script language=javascript>
<!--
function view(user,myself)
{url="sendmess.asp?user="+user+"&myself="+myself;
 privatemsg = window.open(url,"AE","WIDTH=400,HEIGHT=120,left=100,top=250,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,menu=no,min=no,max=no");
 }
-->
</script>
</head>
<body aLink="#000000" bgcolor=#99aacc link="#000000" text="#000000" topMargin="4" Link="#000000" MARGINWIDTH="5" MARGINHEIGHT="4">
<%
'Response.Write request("message")&"whc"
temp=left(request("message"),len(request("message"))-2)
message1=split(temp,"||")
'session("message")=""%>
<div align=center style=font-size:9pt;>点击发信息者的名字给其回信</div>
<TABLE border=1 cellPadding=1 cellSpacing=0 width=90% bordercolor=mediumaquamarine bordercolordark=#aaeed4 bordercolorlight=#404040 align=center>
<tr><td align=center>发信者</td><td align=center width=80%>内容</td></tr>
<%for i=0 to ubound(message1)
message=split(message1(i),"$$")%>
<tr><td align=center><a href=javascript:view('<%=message(0)%>','<%=message(1)%>')><%=message(0)%></a></td><td><%=message(2)%></td></tr>
<%next%>
</table>
<br>
</body>
</html>