<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="std.asp" -->
<%
	OnlyLogin
	Alive "�����ķ�"
%>
<html>
<head>
<script language=javascript>
<!--
function view(user,myself)
{url="sendmess.asp?user="+user+"&myself="+myself;
 privatemsg = window.open(url,"AE","WIDTH=400,HEIGHT=120,left=100,top=250,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,menu=no,min=no,max=no");
 }
-->
</script>
<meta http-equiv="pragma" content="no-cache" >
<link rel="stylesheet" type="text/css" href="main.css">
</head>
<body class=clblue>
<%		dim strSQL,iUser,i,byOrder,deleteid,misspassid

%>
<%
dim username
username=session("userid")
%>
<center>
<%=username%>���!������ֱ�ӵ㴫�����û����Ͷ���Ϣ!
<table width=70% class=main cellspacing=2 align=center >
	<tr class=usertitle>
		<td width='10%' ><center>����</center></td>
		<td width='20%' ><center>�û�ID</center></td>
		<td width='30%' ><center>�س�</center></td>
		<td width='10%' ><center>��վ</center></td>
		<td width='10%' ><center>����</center></td>
		<td width='20%' ><center>״��</center></td>
	</tr>
<%	dim UserList,uT,sT
	UserList = Application("UserList")
	for i =1 to  100
	  if UserList(i,0)<>"" then
		uT=clng((Timer()-UserList(i,1))/60)
		if uT<0 then 
			uT=clng((Timer()/60+1440-UserList(i,1)/60))
		end if
		sT=clng((Timer()-UserList(i,2))/60)
		if sT<0 then 
			sT=clng((Timer()/60+1440-UserList(i,2)/60))
		end if

%> 
		<tr class=userline1>
		<td align="center">
<%if UserList(i,0)<>"guest" then%>
<a href="javascript:view('<%=UserList(i,0)%>','<%=session("userid")%>')">����</a>
<%else%>����
<%end if%>
		</td>
		<td>&nbsp;<a class=UserOnlineA href=queryUser.asp?userid=<%=UserList(i,0)%>><%=UserList(i,0)%></a></td>
		<td><%=UserList(i,3)%></td>
		<td align=center><%=uT%>&nbsp;</td>
		<td align=center><%if sT >0 then %><%=sT%><%end if%>&nbsp;</td>
		<td align=center><%=UserList(i,4)%> </td>
	</tr>
<%	  end if
	next 
%>
</table>
<p>
</body>
</html>

