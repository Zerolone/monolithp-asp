<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Links_ID
	Links_ID = Request("Links_ID")
	CheckNum Links_ID
	
	Dim Page
	Page	= Request("Page")
	
	Sql = "Delete From [Monolith_Links] Where [Links_ID]=" & Links_ID
	Conn.Execute(Sql)
	Call CloseConn

	Response.Redirect "Default.asp?Page=" & Page
%>