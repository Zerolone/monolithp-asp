<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<%
	Dim Tree_Parent_ID, Tree_Level, Tree_Title, Tree_Url, Tree_Target, Tree_HasChild

	Tree_Parent_ID		= Request("Tree_Parent_ID")
	Tree_Level				= Request("Tree_Level")
	Tree_Title				= Request("Tree_Title")
	Tree_Url					= Request("Tree_Url")
	Tree_Target				= Request("Tree_Target")
	Tree_HasChild			= Request("Tree_HasChild")

  Call CheckNum(Tree_Parent_ID)

	Tree_Title				= MonolithEncode(Tree_Title)
	Tree_Url					= MonolithEncode(Tree_Url)

	Sql = "Insert Into [Monolith_Tree] ([Tree_Parent_ID], [Tree_Level], [Tree_Title], [Tree_Url], [Tree_Target], [Tree_HasChild]) values ("
	Sql = Sql & Tree_Parent_ID & ","
	Sql = Sql & "'" & Tree_Level & "',"
	Sql = Sql & "'" & Tree_Title & "',"
	Sql = Sql & "'" & Tree_Url & "',"
	Sql = Sql & "'" & Tree_Target & "',"
	Sql = Sql & "" & Tree_HasChild & ")"

	Conn.execute(Sql)

	Sql	= "Update [Monolith_Tree] Set "
	Sql = Sql & "[Tree_HasChild]=true "
	Sql = Sql & " Where [Tree_ID]=" & Tree_Parent_ID

	Conn.execute(Sql)

	Call CloseConn

	Response.write "<script language=""javascript"">"
	Response.write "  alert(""添加成功, 请刷新页面察看新添加的数据。"");"
	Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
	Response.write "</script>"
%>