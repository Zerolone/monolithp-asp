<% option explicit%>
<!-- #include file="std.asp" -->
<html>
<head>
<script language=javascript>
<!--
function view(message)
{url="showmess.asp?message="+message;
 privatemsg = window.open(url,"AE1","WIDTH=400,HEIGHT=120,left=100,top=100,toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=yes,menu=no,min=no,max=no");
 }
-->
</script>

<meta http-equiv="pragma" content="no-cache">
<meta http-equiv="refresh" content="1500">
<link rel="stylesheet" type="text/css" href="main.css">
</head>
<%
dim messages,message,message1,messagetemp
session("message")=""
if application("message")<>"" then
	messages=left(application("message"),len(application("message"))-2)
	message=split(messages,"||")
	for i=0 to ubound(message)
		message1=split(message(i),"&&")
		if message1(1)=session("userid") then
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
    	Response.write "<body topmargin=0 leftmargin=0 class=clblack onload=view('"&replace(session("message"),"&&","$$")&"')>"
else
    Response.write "<body topmargin=0 leftmargin=0 class=clblack>"

end if%>

<%
'  dim strcounter
'  strcounter=right("00"&cstr(application("Counter")),3)
'  Online:  =strcounter</P>
	dim UserList,i,sT
	UserList = Application("UserList")
	for i=1 to 100
		if UserList(i,0)<>"" then
			sT=clng((Timer()-UserList(i,2))/60)
			if sT<0 then 
				sT=clng((Timer()/60+1440-UserList(i,2)/60))
			end if
			if sT>60 then
				UserList(i,0)=""
				UserList(i,4)="°×³ÕÖÐ"
			end if
			
		end if
	next 
	Application.Lock
	Application("UserList")=UserList
	Application.UnLock
%>
<body leftmargin="0" topmargin="-1">
<div align="center">
</div>
</body>
</html>