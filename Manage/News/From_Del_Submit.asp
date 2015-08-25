<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim From_ID
	From_ID = Request("From_ID")
	CheckNum From_ID
	
	Dim Page
	Page	= Request("Page")	

	Sql = "Delete From [Monolith_News_From] Where [From_ID]=" & From_ID
	Conn.Execute(Sql)
	Call CloseConn

	Response.Redirect "From.asp?Page=" & Page
%>