<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<!-- #include file="std.asp" -->
<%
OnlySysop
dim strSetit,strSSQL
dim rsSysInfo2

strSetit=request("setit")
set rsSysInfo2=CreateObject("ADODB.RecordSet")

if strSetit="����" then
	Application.Lock
	Application("title")=request("title")
    Application.Unlock
	strSSQL="Update sysinfo set title='"+request("title")+"'"
	rsSysInfo2.Open strSSQL,myconn,1
	session("strMsg")="�����������޸ģ���ˢ�»����µ�¼"
end if

if strSetit="����Ա" then
	Application.Lock
	Application("Sysop")=","+request("sysop")+","
	Application.Unlock
	strSSQL="Update sysinfo set sysop='"+Application("sysop")+"'"
	rsSysInfo2.Open strSSQL,myconn,1
	session("strMsg")="����Ա�����Ѿ����£�"
end if

set rsSysInfo2=nothing
session("strURL")="sysManager.asp?o="+cstr(timer())
response.redirect "alert.asp"
%>