<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Advertisement_ID
	Advertisement_ID = Request("Advertisement_ID")
	CheckNum Advertisement_ID
	
	Dim Page
	Page	= Request("Page")	

	Sql = "Delete From [Monolith_Advertisement] Where [Advertisement_ID]=" & Advertisement_ID
	Conn.Execute(Sql)
	Call CloseConn

	Response.Redirect "Default.asp?Page=" & Page

%>