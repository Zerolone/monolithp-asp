<%
if session("username")="" then%>
<script language=javascript>
parent.parent.location="error.asp"

</script>
	
<%else
session("username")=session("username")
session("lasttime")=minute(now())

%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<meta http-equiv="refresh" content="10; URL=menutop.asp">   

<style type="text/css">
A:link {COLOR: #000000; TEXT-DECORATION: none; font-size: 9pt}
A:active {COLOR: #000000; TEXT-DECORATION: none; font-size: 9pt}
A:visited {COLOR: #000000; TEXT-DECORATION: none; font-size: 9pt}
.song9 {  font-family: 宋体; font-size: 9pt}
</style>
<script language=javascript>
<!--
function view(message)
{url="showmess.asp?message="+message;
 privatemsg = window.open(url,"AE1","WIDTH=400,HEIGHT=120,left=100,top=100,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,menu=no,min=no,max=no");
 }
-->
</script>


<%session("message")=""
if application("message")<>"" then
	messages=left(application("message"),len(application("message"))-2)
	message=split(messages,"||")
	for i=0 to ubound(message)
		message1=split(message(i),"&&")
		if message1(1)=session("username") then
		    session("message")=session("message")&message(i)&"||"
		    	
		else
			messagetemp=messagetemp&message(i)&"||"
		end if	
	next
	Application.Lock
	application("message")=messagetemp
	Application.UnLock
end if
if session("message")<>"" then
    	Response.write "<body aLink=#ffee00 bgColor=#6699cc link=#ffffff text=#ffffff topMargin=4 Link=#000000 topmargin=0 leftmargin=0 onload=view('"&replace(session("message"),"&&","$$")&"')>"
	Response.write "<div align=center><a href=onlineuser.asp target=main><font color=ffffff>查看在线人数</a></div>"
else
    Response.write "<body aLink=#ffee00 bgColor=#6699cc link=#ffffff text=#ffffff topMargin=4 Link=#000000 topmargin=0 leftmargin=0>"
    Response.write "<div align=center><a href=onlineuser.asp target=main><font color=ffffff>查看在线人数</a></div>"
	
end if%>
</body>
</html>
<%end if%>