<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Author_ID
	Author_ID = Request("Author_ID")
	CheckNum Author_ID
	
	Dim Page
	Page	= Request("Page")	

	Sql = "Delete From [Monolith_News_Author] Where [Author_ID]=" & Author_ID
	Conn.Execute(Sql)
	Call CloseConn

	Response.Redirect "Author.asp?Page=" & Page
%>