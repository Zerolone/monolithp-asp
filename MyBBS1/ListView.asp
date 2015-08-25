<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="Inc/Skin.asp"-->
<!-- #include file="TopMenu.asp"-->
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<hr>
当前页面：帖子察看
<hr>
<%
  dim BBSBoardID
  dim BoardStyle, Arr_BoardStyle
  dim ID 
  ID			= Request("ID")
  BBSBoardID 	= Request("BBSBoardID")
  sql= "select [BoardName],[BoardStyle] from [MyBBS_Board] Where [ID]=" & BBSBoardID
  rs.open sql, conn, 1, 3
  SqlQueryNum	= SqlQueryNum + 1
  if not (rs.bof or rs.eof) then
'    BoardName 		= Rs(0)
    BoardStyle		= Rs(1)
    Arr_BoardStyle	= Split(BoardStyle,",")
  else
    Response.write "错误的板块！"
	Response.end
  end if
  rs.close
   
  '--------------0----1-----------2--------------3-----------4----------5-------------6----------7
  sql= "select [ID], [BBSTitle], [BBSPostUser], [BBSReply], [BBSView], [BBSPostTime], [BBSBody], [BBSBoardID] from [MyBBS_BBS] Where [ID]=" & ID
'  Response.write Sql
  rs.open sql, conn, 1, 3
  SqlQueryNum	= SqlQueryNum + 1
  if err.number <> 0 then
    response.write "数据库操作失败，请稍候重试！"
	response.end
  else
	if (rs.bof or rs.eof) then
	  Response.write "参数错误！"
	  Response.end
    else
      dim BBSTitle, BBSPostUser, BBSReply, BBSView, BBSPostTime, BBSBody
	  Rs(4) = Rs(4) + 1
	  ID					= Rs(0)
	  BBSTitle				= Rs(1)
	  BBSPostUser			= Rs(2)
	  BBSReply				= Rs(3)
 	  BBSView				= Rs(4)
	  BBSPostTime			= Rs(5)
	  BBSBody				= Rs(6)
	  BBSBoardID			= Rs(7)
	  Rs.Update
	  Rs.close
	  
	  Dim ListName 
	  ListName= BBSTitle
	  
	  
	  sql= "select [ID], [UserName], [Topic] from [MyBBS_Users] Where [UserName]='" & BBSPostUser & "'"
	  'Response.write Sql
	  rs.open sql, conn, 1, 3
	  SqlQueryNum	= SqlQueryNum + 1
	  if err.number <> 0 then
	    response.write "数据库操作失败，请稍候重试！"
		response.end
	  else
	    dim ID_Users, UserName_Users, Topic_Users

	    if (rs.bof or rs.eof) then
		  ID_Users			= "该用户不存在或者被管理员屏蔽"
		  UserName_Users	= "该用户不存在或者被管理员屏蔽"
		  Topic_Users		= "该用户不存在或者被管理员屏蔽"
		else
		  ID_Users			= Rs(0)
		  UserName_Users	= Rs(1)
		  Topic_Users		= Rs(2)
		end if
		Rs.close
	  end if
 
	  if IsNull(BBSPostTime) 			then	BBSPostTime		= "　"

	  TheString = ""
	  TheString = Replace(Skin_ListViewMain, "{BBSTitle}", BBSTitle)
	  TheString = Replace(TheString, "{StyleBgImage}"	, Trim(Arr_BoardStyle(0)))
	  TheString = Replace(TheString, "{StyleTr1}"		, Trim(Arr_BoardStyle(1)))
	  TheString = Replace(TheString, "{StyleFoot}"		, Trim(Arr_BoardStyle(3)))
	  TheString = Replace(TheString, "{ID}"				, ID_Users)
	  TheString = Replace(TheString, "{UserName}"		, UserName_Users)
	  TheString = Replace(TheString, "{Topic}"			, Topic_Users)
	  TheString = Replace(TheString, "{BBSBody}"		, BBSBody)
	  TheString = Replace(TheString, "{BBSReply}"		, BBSReply)
	  TheString = Replace(TheString, "{BBSView}"		, BBSView)
	  TheString = Replace(TheString, "{BBSPostTime}"	, BBSPostTime)

	  'TheString = Replace(TheString, "{BoardID}", ID)
	  Response.write TheString +  Chr(10)
    end if
  end if

  Response.write "<hr>以下为回复内容：<br>"
  '--------------0------------1--------------2------------3		
  Sql = "Select [BBSTitle], [BBSPostUser], [BBSPostTime],[BBSBody] From [MyBBS_BBS] Where [BBSRoot]=" & ID
  Sql = Sql & " And BBSBoardID=" & BBSBoardID
  Sql = Sql & " Order By [ID] Asc"
  Rs.open Sql, Conn ,1, 3
  SqlQueryNum	= SqlQueryNum + 1
  if err.number <> 0 then
    response.write "数据库操作失败，请稍候重试！"
	response.end
  else
    Response.write "<script language=javascript>" & Chr(10)
    '-----------分页
    dim page, news, PageSize, RsNo, pgnum, Count, ThePageUrl, ThisPage
    dim RsTemp
    set RsTemp		= Server.CreateObject("ADODB.Recordset")

    ThisPage	= Request.ServerVariables("PATH_INFO")
    Count = 0
    page=request("page")
    PageSize = 5
    Rs.PageSize = PageSize
    RsNo=Rs.recordcount
    pgnum=Rs.Pagecount
    if page="" or clng(page)<1 then page=1
    if clng(page) > pgnum then page=pgnum
    if pgnum>0 then Rs.AbsolutePage=page						
    '=-=-=--=-=-  

    do while not (rs.bof or rs.eof) and count<Rs.PageSize
	  BBSTitle				= Rs(0)
	  BBSPostUser			= Rs(1)
	  BBSPostTime			= Rs(2)
	  BBSBody				= Rs(3)

      sql= "select [ID], [UserName], [Topic] from [MyBBS_Users] Where [UserName]='" & BBSPostUser & "'"
	  'Response.write Sql
	  rsTemp.open sql, conn, 1, 3
	  SqlQueryNum	= SqlQueryNum + 1
	  if err.number <> 0 then
	    response.write "数据库操作失败，请稍候重试！"
	    Response.end
	  else
	    if (rsTemp.bof or rsTemp.eof) then
		  ID_Users				= "该用户不存在或者被管理员屏蔽"
  		  UserName_Users		= "该用户不存在或者被管理员屏蔽"
		  Topic_Users			= "该用户不存在或者被管理员屏蔽"
	    else
		  ID_Users				= RsTemp(0)
		  UserName_Users		= RsTemp(1)
		  Topic_Users			= RsTemp(2)
	    end if
	    RsTemp.close
	  end if	
	
	  TheString = ""
	  TheString = Replace(Skin_ListViewSecond, "{BBSTitle}", BBSTitle)
	  TheString = Replace(TheString, "{ID}"				, ID_Users)
	  TheString = Replace(TheString, "{UserName}"		, UserName_Users)
	  TheString = Replace(TheString, "{Topic}"			, Topic_Users)
	  TheString = Replace(TheString, "{BBSBody}"		, BBSBody)
	  TheString = Replace(TheString, "{BBSReply}"		, BBSReply)
	  TheString = Replace(TheString, "{BBSView}"		, BBSView)
	  TheString = Replace(TheString, "{BBSPostTime}"	, BBSPostTime)
	  Response.write "MyBBS_List.innerHTML=MyBBS_List.innerHTML + """ & TheString & """;"+  Chr(10)

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
  
  Response.write "<hr>[<font color=red>" & ListName & "</font>] 合计<font color=red><b>" & RsNo & "</b></font>条  |"
  ThePageUrl	= ThisPage & "?ID=" & ID & "&BBSBoardID=" & BBSBoardID 
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

  Response.write "<hr>"
  if Username <> "" then
%>
<form action="ListSave.asp" method="post" name="ReplyTopic">
  <p>回复主题： 
    <input name="BBSTitle" type="text" size="51">
    <br>
    回复内容： 
    <textarea name="BBSBody" cols="50" rows="5"></textarea>
    <input type="hidden" name="BBSBoardID" value="<%=BBSBoardID%>"><input type="hidden" name="BBSRoot" value="<%=ID%>">
    <br>
    <input type="submit" name="Submit" value=" 发 表 我 的 回 复 ">
  </p>
</form>
<% end if %>
<!-- #include file="BoardTime.asp"-->