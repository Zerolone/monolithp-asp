<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<%
	Dim strData
	Dim intFieldCount
	Dim i

	intFieldCount = Request.Form("hdnCount")
	For i=1 To intFieldCount
		strData = strData & Request.Form("hdnTemplate_Content" & i)
	Next
	
	Dim Template_Parent_ID, Template_Level, Template_Title, Template_HasChild, Template_Content

	Template_Parent_ID		= Request("Template_Parent_ID")
	Template_Level				= Request("Template_Level")
	Template_Title				= Request("Template_Title")
	Template_HasChild			= Request("Template_HasChild")
	Template_Content			= strData

  Call CheckNum(Template_Parent_ID)

	Template_Title				= MonolithEncode(Template_Title)

	'-----------------------0-------------------1----------------2----------------3----------------------4
	Sql = "Select [Template_Parent_ID], [Template_Level], [Template_Title], [Template_HasChild], [Template_Content] From [Monolith_Template]"
	Rs.open Sql, conn , 1, 3
	Rs.AddNew
	Rs(0) = Template_Parent_ID
	Rs(1) = Template_Level
	Rs(2) = Template_Title
	Rs(3)	= Template_HasChild
	Rs(4)	= Template_Content
	Rs.Update

	Sql	= "Update [Monolith_Template] Set "
	Sql = Sql & "[Template_HasChild]=true "
	Sql = Sql & " Where [Template_ID]=" & Template_Parent_ID

	Conn.execute(Sql)

	Call CloseRs
	Call CloseConn

	Response.write "<script language=""javascript"">"
	Response.write "  alert(""添加成功, 请刷新页面察看新添加的数据。"");"
	Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
	Response.write "</script>"
%>