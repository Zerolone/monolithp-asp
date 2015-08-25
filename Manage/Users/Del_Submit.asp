<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Users_ID
	Users_ID = Request("Users_ID")
	CheckNum Users_ID
	
	Dim Page
	Page	= Request("Page")	

	Sql = "Delete From [Monolith_Users] Where [Users_ID]=" & Users_ID
	Conn.Execute(Sql)
	Call CloseConn

	Response.Redirect "Default.asp?Page=" & Page
%>