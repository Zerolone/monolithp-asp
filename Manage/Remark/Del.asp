<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Remark_ID
	Remark_ID = Request("Remark_ID")
	CheckNum Remark_ID

	Dim News_ID
	News_ID = Request("News_ID")
	CheckNum News_ID
	
	Dim Page
	Page		= Request("Page")

	Sql = "Delete From [Monolith_News_Remark] Where [Remark_ID]=" & Remark_ID
	Conn.Execute(Sql)
	
	Sql	= "Update [Monolith_News] Set [News_Remark]= [News_Remark]-1 Where [News_ID]=" & News_ID
	Conn.Execute(Sql)


	Call CloseConn
	Response.Redirect "Default.asp?Page=" & Page
%>