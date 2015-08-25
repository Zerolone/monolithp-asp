<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!--#include virtual="/Manage/Check.asp"-->
<script language="javascript" src="/Js/All.js"></script>
<script language="javascript" src="/Js/MdiWin.js"></script>
<%
	Dim Bulletin_ID, Bulletin_Title, Bulletin_Title2, Bulletin_Content, Bulletin_Content2

	Bulletin_Title			= Request("Bulletin_Title")
	Bulletin_Title2			= Request("Bulletin_Title2")
	Bulletin_Content		= Request("Bulletin_Content")
	Bulletin_Content2		= Request("Bulletin_Content2")

	Bulletin_Title			= MonolithEncode(Bulletin_Title)
	Bulletin_Title2			= MonolithEncode(Bulletin_Title2)
	Bulletin_Content		= MonolithEncode(Bulletin_Content)
	Bulletin_Content2		= MonolithEncode(Bulletin_Content2)
	
	'----------------------0-----------------1------------------2-------------------3--------------------4
	Sql = "Select [Bulletin_Title], [Bulletin_Title2], [Bulletin_Content], [Bulletin_Content2], [Bulletin_ID]  From [Monolith_Bulletin]"
	Rs.Open Sql, Conn, 1, 3
	Rs.AddNew
	Rs(0)								= Bulletin_Title
	Rs(1)								= Bulletin_Title2
	Rs(2)								= Bulletin_Content
	Rs(3)								= Bulletin_Content2
	
	Bulletin_ID					= Rs(4)
	Rs.Update
	Call CloseRs
	Call CloseConn
	
	Response.write "<script language=""javascript"">"
	Response.write "  alert(""添加成功！请刷新公告列表"");"
	Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
	Response.write "</script>"	
%>