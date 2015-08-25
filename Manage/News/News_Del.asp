<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<%
	dim News_ID, Page, News_Class
	News_ID			= Request("News_ID")
	Page				= Request("Page")
	News_Class	= Request("News_Class")
	CheckNum News_ID
	Sql="DELETE FROM [Monolith_News] where [News_ID]="& News_ID
	Conn.Execute (Sql)
	Call CloseConn
	Response.Redirect "News_List.asp?Page=" & Page & "&News_Class=" + News_Class
%>