<!-- #include virtual =	"/Inc/Monolith_Conn.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<%
	'---网站字符串----临时字符串1---临时字符串1
	Dim TheString,    TempString1,  TempString2 

	TheString		= ""
	'调首页
	TheString		= TheString & LoadTemplate("04_03_GuestBook")

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

	Dim ThisPage
  ThisPage				= Request.ServerVariables("PATH_INFO")
  TheString		= Replace(TheString, "{ThisPage}", ThisPage)

	TempString1	= Split(TheString, "<!--Split-->")
	TempString2	= ""


	'---------------------0-----------------------1-----------------2--------------------3
	Sql="Select [GuestBook_Users_Name], [GuestBook_Mail], [GuestBook_Content], [GuestBook_Postdate], [GuestBook_Replay], [GuestBook_ReplayDate] from [Monolith_GuestBook] order by [GuestBook_Postdate] desc"
	Rs.open Sql, Conn, 1, 3
	if Request("send")="ok" then
		if Request("GuestBook_Users_Name")="" or Request("GuestBook_Content")="" then
			response.write "<script language='javascript'>"
			response.write "alert('无效输入！请检查您填写的内容');"
			response.write "location.href='" & ThisPage & "';"			
			response.write "</script>"
		else
			Rs.Addnew
			Rs(0)	= Request("GuestBook_Users_Name")
			Rs(1)	= Request("GuestBook_Mail")
			Rs(2)	= Request("GuestBook_Content")
			Rs(3)	= Now()
			Rs.Update
	    Response.Redirect ThisPage
		end if
	end if

	Do While Not(Rs.Bof Or Rs.Eof)
		TempString2	= TempString2 & TempString1(1)
		TempString2	= Replace(TempString2, "{GuestBook_Users_Name}" , 	Rs(0))
		TempString2	= Replace(TempString2, "{GuestBook_Mail}" , 				Rs(1))
		TempString2	= Replace(TempString2, "{GuestBook_Content}" , 			Rs(2))
		TempString2	= Replace(TempString2, "{GuestBook_Postdate}" ,	 		Rs(3))
   	Rs.MoveNext
  Loop

  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1

	TheString		= TempString1(0) & TempString2 & TempString1(2)

	Response.write TheString

	Call CloseConn
%>