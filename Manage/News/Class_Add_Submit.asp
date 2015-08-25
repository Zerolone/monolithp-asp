<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<%
	Dim News_Class_Parent_ID, News_Class_Level, News_Class_Title, News_Class_HasChild

	News_Class_Parent_ID		= Request("News_Class_Parent_ID")
	News_Class_Level				= Request("News_Class_Level")
	News_Class_Title				= Request("News_Class_Title")
	News_Class_HasChild			= Request("News_Class_HasChild")

  Call CheckNum(News_Class_Parent_ID)

	News_Class_Title				= MonolithEncode(News_Class_Title)

	Sql = "Insert Into [Monolith_News_Class] ([News_Class_Parent_ID], [News_Class_Level], [News_Class_Title], [News_Class_HasChild]) values ("
	Sql = Sql & News_Class_Parent_ID & ","
	Sql = Sql & "'" & News_Class_Level & "',"
	Sql = Sql & "'" & News_Class_Title & "',"
	Sql = Sql & "" & News_Class_HasChild & ")"

	Conn.execute(Sql)

	Sql	= "Update [Monolith_News_Class] Set "
	Sql = Sql & "[News_Class_HasChild]=true "
	Sql = Sql & " Where [News_Class_ID]=" & News_Class_Parent_ID

	Conn.execute(Sql)

	Call CloseConn

	Response.Redirect "Class_Setting.asp"

%>