<!-- #include virtual =	"/Inc/Monolith_Conn.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<%
	Dim Bulletin_ID
	Bulletin_ID = Request("Bulletin_ID")
	CheckNum Bulletin_ID

	'---网站字符串----临时字符串1---临时字符串1
	Dim TheString,    TempString1,  TempString2 

	TheString		= ""
	'调首页
	TheString		= TheString & LoadTemplate("04_02_Bulletin")

	'调头
	TheString		= Replace(TheString, "{Monolith_Head}", LoadTemplate("02_02_Default_Head"))

	'Css
	TheString		= Replace(TheString, "{Monolith_Css}", 	LoadTemplate("01_01_Css"))

	'Foot
	TempString1	= LoadTemplate("01_02_Foot")
  TempString1	= Replace(TempString1, "{Now}",	Now())
  TempString1	= Replace(TempString1, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "毫秒。")
  TempString1	= Replace(TempString1, "{SqlQueryNum}",	SqlQueryNum)
	
	TheString		= Replace(TheString, "{Monolith_Foot}", 	TempString1)

	'分隔线
	TheString		= Replace(TheString, "{Monolith_Sep}", 	LoadTemplate("01_03_Sep"))


	'----------------------0-----------------1-------------------2
	Sql = "Select [Bulletin_Title], [Bulletin_Content], [Bulletin_Date] From [Monolith_Bulletin] Where [Bulletin_ID]=" & Bulletin_ID
	Rs.Open Sql, Conn, 1, 1
	if Rs.Bof Or Rs.Eof then
		Response.write "公告不存在！"
		Response.end
	else
		TheString	= Replace(TheString, "{Bulletin_Title}" , 	Rs(0))
		TheString	= Replace(TheString, "{Bulletin_Content}" , Replace(Rs(1), chr(10) , "<br>"))
		TheString	= Replace(TheString, "{Bulletin_Date}" , 		Rs(2))
	end if

	Response.write TheString

	Call CloseConn
%>