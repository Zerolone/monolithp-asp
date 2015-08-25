<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<%
	Dim Info_Class_Parent_ID, Info_Class_Level, Info_Class_Title, Info_Class_HasChild

	Info_Class_Parent_ID		= Request("Info_Class_Parent_ID")
	Info_Class_Level				= Request("Info_Class_Level")
	Info_Class_Title				= Request("Info_Class_Title")
	Info_Class_HasChild			= Request("Info_Class_HasChild")

  Call CheckNum(Info_Class_Parent_ID)

	Info_Class_Title				= MonolithEncode(Info_Class_Title)

	Sql = "Insert Into [Monolith_Info_Class] ([Info_Class_Parent_ID], [Info_Class_Level], [Info_Class_Title], [Info_Class_HasChild]) values ("
	Sql = Sql & Info_Class_Parent_ID & ","
	Sql = Sql & "'" & Info_Class_Level & "',"
	Sql = Sql & "'" & Info_Class_Title & "',"
	Sql = Sql & "" & Info_Class_HasChild & ")"

	Conn.execute(Sql)

	Sql	= "Update [Monolith_Info_Class] Set "
	Sql = Sql & "[Info_Class_HasChild]=true "
	Sql = Sql & " Where [Info_Class_ID]=" & Info_Class_Parent_ID

	Conn.execute(Sql)

	Call CloseConn

	Response.Redirect "Class_Setting.asp"

%>