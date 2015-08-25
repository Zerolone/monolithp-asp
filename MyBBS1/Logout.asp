<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/Config.asp"-->
<%
  Response.cookies(MyBBSCookiesName)("UserName")	= ""
  Response.cookies(MyBBSCookiesName)("Password")	= ""
  Response.cookies(MyBBSCookiesName).expires		= date
  IsLogin											= 0
  Response.redirect "index.asp"
%>
<!-- #include file="BoardTime.asp"-->