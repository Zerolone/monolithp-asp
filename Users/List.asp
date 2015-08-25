<!-- #include virtual =	"/Inc/Monolith_Conn.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<%
	'---网站字符串----临时字符串1---临时字符串1
	Dim TheString,    TempString1,  TempString2 

	TheString		= ""
	'调首页
	TheString		= TheString & LoadTemplate("05_02_UserList")

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


	'------------------
	TempString1	= Split(TheString, "<!--Split-->")
	TempString2	= ""
	
	
	Dim Page, Pages, j, ThisPage
  ThisPage							= Request.ServerVariables("PATH_INFO")
	Page									= clng(request("Page"))

  '------------------0-------------1--------------2--------------3----------------4----------------5
  Sql= "Select [Users_Name], [Users_Level], [Users_Group], [Users_RegTime], [Users_Subject], [Users_Email] From [Monolith_Users] Order By [Users_ID] Desc"

  Dim i 
  i = 1
  Dim Page_Break

  Rs.open Sql, Conn, 1, 1
	if Not (Rs.bof Or Rs.Eof) then
		Rs.PageSize=10
		if page=0 then page=1
		pages=rs.pagecount
		if page > pages then page=pages
		rs.AbsolutePage=page
		for j=1 to rs.PageSize 
	  	TempString2	= TempString2 & TempString1(1)
			TempString2 = Replace(TempString2 , "{Users_Name}",				Rs(0))
			TempString2 = Replace(TempString2 , "{Users_Level}",			Rs(1))
			TempString2 = Replace(TempString2 , "{Users_Group}",			Rs(2))
			TempString2 = Replace(TempString2 , "{Users_RegTime}",		Rs(3))
			TempString2 = Replace(TempString2 , "{Users_Subject}",		Rs(4))
			TempString2 = Replace(TempString2 , "{Users_Email}",			Rs(5))

			i = i + 1
    	Rs.MoveNext
    	if rs.eof then exit for
  	next

		'Page Break
		Page_Break = "<table  cellspacing='0' width='98%' align='center' height='26' cellpadding='0'>"
		Page_Break = Page_Break & "<form method='Post' action='" & ThisPage & "'>"
		Page_Break = Page_Break & "<tr>"
		Page_Break = Page_Break & "<td align=left>"

	  ThisPage = ThisPage & "?1=1"
	  if Page<2 then
	  	Page_Break = Page_Break & "｜&nbsp;<strike>最前页</strike>&nbsp;｜<strike>上一页</strike>&nbsp;｜"
		else
	  	Page_Break = Page_Break & "｜&nbsp;<a href='" & ThisPage & "&page=1'>最前页</a>&nbsp;｜<a href='" & ThisPage & "&page=" & Page-1 & "'>上一页</a>&nbsp;｜"			
	  end if
		if rs.pagecount-page<1 then
			Page_Break = Page_Break & "<strike>下一页</strike> ｜<strike>最后页</strike> ｜"			
		else
			Page_Break = Page_Break & "<a href='" & ThisPage & "&page=" & page+1 & "'>下一页</a>&nbsp;｜<a href='" & ThisPage & "&page=" & rs.pagecount & "'>最后页</a>｜"
		end if
		Page_Break = Page_Break & "页次：<b><font color=red>" & Page & "</font>/" & rs.pagecount & "</b>&nbsp;页&nbsp;｜"
		Page_Break = Page_Break & "共 <b><font color='#FF0000'>" & rs.recordcount & "</font></b> 个回复&nbsp;｜"
		Page_Break = Page_Break & "<b>" & rs.pagesize & "</b> 个回复/页&nbsp;｜"
		Page_Break = Page_Break & "转到：<input type='text' name='page' size=4 maxlength=10 class='InputBox' value='" & page & "'> <input class='InputBox' type='submit'  value=' Goto '  name='cndok'>"
		Page_Break = Page_Break & "</td>"
		Page_Break = Page_Break & "</tr>"
		Page_Break = Page_Break & "</form>"
		Page_Break = Page_Break & "</table>"
	end if

  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1

	TheString		= TempString1(0) & TempString2 & TempString1(2)

	Response.write TheString

	Call CloseConn

	Response.write "<script language='javascript'>"
	Response.write "Page_Break_Head.innerHTML=""" & Page_Break & """" & ";"
	Response.write "Page_Break_Foot.innerHTML=""" & Replace(Page_Break, "left", "right") & """" & ";"
	Response.write "</script>"
%>