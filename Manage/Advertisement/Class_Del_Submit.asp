<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Advertisement_Class_ID, Cmd
	Advertisement_Class_ID = Request("Advertisement_Class_ID")
	CheckNum Advertisement_Class_ID
	
	Dim Page
	Page	= Request("Page")
	
	Sql = "Delete From [Monolith_Advertisement_Class] Where [Advertisement_Class_ID]=" & Advertisement_Class_ID
	Conn.Execute(Sql)
	Call CloseConn

	Response.Redirect "Class.asp?Page=" & Page
%>