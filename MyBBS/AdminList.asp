<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="Inc/Skin.asp"-->
<!-- #include file="TopMenu.asp"-->
<!-- #include file="Inc/AdminJs.asp"-->
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<hr>
当前页面：帖子管理列表
<hr>
<%
  dim BoardID
  
  dim BoardName , BoardStyle, Arr_BoardStyle, ThisPage

  BoardID 	= Request("BoardID")
  ThisPage	= Request.ServerVariables("PATH_INFO")

  sql= "select [BoardName],[BoardStyle] from [MyBBS_Board] Where [ID]=" & BoardID
  rs.open sql, conn, 1, 3
  SqlQueryNum	= SqlQueryNum + 1
  if not (rs.bof or rs.eof) then
    BoardName 		= Rs(0)
    BoardStyle		= Rs(1)
    Arr_BoardStyle	= Split(BoardStyle,",")
  else
    Response.write "错误的板块！"
	Response.end
  end if
  rs.close

  if err.number <> 0 then
    response.write "数据库操作失败，请稍候重试！"
	response.end
  else
    TheString = ""
    TheString = Replace(Skin_ListMain, "{BoardName}", BoardName)
	TheString = Replace(TheString, "{StyleBgImage}"	, Trim(Arr_BoardStyle(0)))
	TheString = Replace(TheString, "{StyleTr1}"		, Trim(Arr_BoardStyle(1)))
	TheString = Replace(TheString, "{StyleFoot}"	, Trim(Arr_BoardStyle(3)))

'    TheString = Replace(TheString, "{BoardID}", ID)
    Response.write "<a href='ListAdd.asp?BoardID=" & BoardID & "'>发帖</a>"
    Response.write TheString +  Chr(10)
  end if

  TheString = ""
  '-------------0-----1------------2--------------3-------------4-------------5--------------6-------------7-------------8--------------------9--------------------10-------------------
  sql= "select [ID], [BBSTitle], [BBSPostUser], [BBSReply], [BBSView], [BBSLastPostTime], [BBSLastPostUser] from [MyBBS_BBS] Where [BBSBoardID]=" & BoardID & " and [BBSRoot]=0 order by [ID] DESC"
  rs.open sql, conn, 1, 3
  SqlQueryNum	= SqlQueryNum + 1
  if err.number <> 0 then
    response.write "数据库操作失败，请稍候重试！"
	response.end
  else
    Response.write "<script language=javascript>" & Chr(10)
	if (rs.bof or rs.eof) then Response.write "MyBBS_List.innerHTML=MyBBS_List.innerHTML + ""当前板块没有帖子，你可以--<a href='ListAdd.asp?BoardID=" & BoardID & "'>发表你的帖子!</a>;"""+  Chr(10)

    dim ID, BBSTitle, BBSPostUser, BBSReply, BBSView, BBSLastPostTime, BBSLastPostUser
	
	'-----------分页
	dim page, news, PageSize, RsNo, pgnum, Count, ThePageUrl
	Count = 0
	page=request("page")
	PageSize = 15						
	Rs.PageSize = PageSize
	RsNo=Rs.recordcount
	pgnum=Rs.Pagecount
	if page="" or clng(page)<1 then page=1
	if clng(page) > pgnum then page=pgnum
	if pgnum>0 then Rs.AbsolutePage=page						
	'=-=-=--=-=-
	do while not (rs.bof or rs.eof) and count<Rs.PageSize
	  ID					= Rs(0)
	  BBSTitle				= Rs(1)
	  BBSPostUser			= Rs(2)
	  BBSReply				= Rs(3)
	  BBSView				= Rs(4)
	  BBSLastPostTime		= Rs(5)
	  BBSLastPostUser		= Rs(6)

	  if IsNull(BBSLastPostTime) 		then	BBSLastPostTime		= "　"
	  if IsNull(BBSLastPostUser) 		then	BBSLastPostUser		= "　"

	  TheString = ""
	  TheString = Replace(Skin_ListSecond, "{BBSTitle}"	,	"｜<a  onclick='return cformdeltopic();' href='AdminListAction.asp?Code=Topic&ID=" & ID & "&BoardID=" & BoardID & "'>删除帖子</a>｜<a href='AdminListView.asp?ID=" & ID & "'>" & BBSTitle & "</a>")
	  TheString = Replace(TheString, "{BBSPostUser}"	,	BBSPostUser)
	  TheString = Replace(TheString, "{BBSReply}"		,	BBSReply)
	  TheString = Replace(TheString, "{BBSView}"		,	BBSView)
	  TheString = Replace(TheString, "{BBSLastPostTime}",	BBSLastPostTime)
	  TheString = Replace(TheString, "{BBSLastPostUser}",	BBSLastPostUser)
	  Response.write "MyBBS_List.innerHTML=MyBBS_List.innerHTML + """ & TheString & """;"+  Chr(10)
	  'Response.Flush()
	  Count = Count + 1
	  rs.movenext
	loop
	
	dim JsString 
	JsString = "var j=0;" + Chr(10)
  	JsString = JsString + "for(var i=0;i<document.all.length;i++){" + Chr(10)
	JsString = JsString + "if (document.all[i].name=='tablelist'){" + Chr(10)
	JsString = JsString + "  if (j % 2 == 0)"
	JsString = JsString + "  {document.all[i].bgColor='" & Trim(Arr_BoardStyle(2)) & "';}" + Chr(10)
	JsString = JsString + "  else" + Chr(10)
	JsString = JsString + "  {document.all[i].bgColor='" & Trim(Arr_BoardStyle(1)) & "';}" + Chr(10)
	JsString = JsString + "  j=j+1" + Chr(10)
	JsString = JsString + " }" + Chr(10)
	JsString = JsString + "}" + Chr(10)

	Response.write JsString
    Response.write "</script>" & Chr(10)
  end if

  Response.write "[<font color=red>" & BoardName & "</font>] 合计<font color=red><b>" & RsNo & "</b></font>条  |"
  ThePageUrl	= ThisPage & "?BoardID=" & BoardID
  if page <=1 then
	Response.write "  上一页 "
  else
	Response.write "  <a href=""" & ThePageUrl & "&page=" & page-1 & """>上一页</a> "
  end if

  if Rs.pagecount-page<1 then
	Response.write "  下一页 "
  else
	Response.write "  <a href=""" & ThePageUrl & "&page=" & page+1 & """>下一页</a> "
  end if
  Response.write "| 页次：<b><font color=red> " & page & "</font>/" & Rs.pagecount & "</b>页 <b>" & PageSize & "</b>条/页 "
  rs.close
%>
<!-- #include file="BoardTime.asp"-->