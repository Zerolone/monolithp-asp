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

if strSetit="改名" then
	Application.Lock
	Application("title")=request("title")
    Application.Unlock
	strSSQL="Update sysinfo set title='"+request("title")+"'"
	rsSysInfo2.Open strSSQL,myconn,1
	session("strMsg")="社区名字已修改，请刷新或重新登录"
end if

if strSetit="管理员" then
	Application.Lock
	Application("Sysop")=","+request("sysop")+","
	Application.Unlock
	strSSQL="Update sysinfo set sysop='"+Application("sysop")+"'"
	rsSysInfo2.Open strSSQL,myconn,1
	session("strMsg")="管理员名单已经更新！"
end if

set rsSysInfo2=nothing
session("strURL")="sysManager.asp?o="+cstr(timer())
response.redirect "alert.asp"
%>