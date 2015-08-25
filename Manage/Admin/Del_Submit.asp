<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Admin_ID
	Admin_ID = Request("Admin_ID")
	CheckNum Admin_ID
	
	Dim Page
	Page	= Request("Page")
	
	Sql = "Delete From [Monolith_Admin] Where [Admin_ID]=" & Admin_ID
	Conn.Execute(Sql)
	Call CloseConn

	Response.Redirect "Default.asp?Page=" & Page
%>