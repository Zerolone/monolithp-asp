<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="Inc/Skin.asp"-->
<!-- #include file="TopMenu.asp"-->
<!-- #include file="Inc/AdminJs.asp"-->

<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<hr>
当前页面：帖子察看
<hr>
<%
  dim id 
  id	= Request("ID")
  '--------------0----1-----------2--------------3-----------4----------5------------------6------------------7---------8
  sql= "select [ID], [BBSTitle], [BBSPostUser], [BBSReply], [BBSView], [BBSLastPostTime], [BBSLastPostUser], [BBSBody], [BBSBoardID] from [MyBBS_BBS] Where [ID]=" & ID
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
      dim BBSTitle, BBSPostUser, BBSReply, BBSView, BBSLastPostTime, BBSLastPostUser, BBSBody, BBSBoardID
	  Rs(4) = Rs(4) + 1
	  ID					= Rs(0)
	  BBSTitle				= Rs(1)
	  BBSPostUser			= Rs(2)
	  BBSReply				= Rs(3)
 	  BBSView				= Rs(4)
	  BBSLastPostTime		= Rs(5)
	  BBSLastPostUser		= Rs(6)
	  BBSBody				= Rs(7)
	  BBSBoardID			= Rs(8)
	  Rs.Update
	  
	  if IsNull(BBSLastPostTime) 		then	BBSLastPostTime		= "　"
	  if IsNull(BBSLastPostUser) 		then	BBSLastPostUser		= "　"
	  
	  TheString = "帖子ID：" & ID
	  TheString = TheString + "<HR>标题：" & BBSTitle
	  TheString = TheString + "<br>内容：" & BBSBody
	  TheString = TheString + "<br>发表人：" & BBSPostUser
	  TheString = TheString + "<br>回复数：" & BBSReply
	  TheString = TheString + "<br>点击数：" & BBSView
	  TheString = TheString + "<br>回复时间：" & BBSLastPostTime
	  TheString = TheString + "<br>回复人：" & BBSLastPostUser 
	  Response.write TheString
    end if
    rs.close
  end if
  Response.write "<hr>以下为回复内容：<br>"
  
  Sql = "Select [BBSTitle], [BBSPostUser], [BBSLastPostTime], [BBSLastPostUser], [BBSBody], [ID] From [MyBBS_BBS] Where [BBSRoot]=" & ID
  Sql = Sql & " And BBSBoardID=" & BBSBoardID
  Sql = Sql & " Order By [ID] Asc"
  Rs.open Sql, Conn ,1, 3

  '-----------分页
  dim page, news, PageSize, RsNo, pgnum, Count, ThePageUrl, ThisPage
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
	BBSTitle			= Rs(0)
	BBSPostUser			= Rs(1)
	BBSLastPostTime		= Rs(2)
	BBSLastPostUser		= Rs(3)
	BBSBody				= Rs(4)
	
	TheString = "<HR>标题：" & BBSTitle
	TheString = TheString + "<bR>内容：" & BBSBody
	TheString = TheString + "<bR>发表人：" & BBSPostUser
	TheString = TheString + "<bR>回复时间：" & BBSLastPostTime
	TheString = TheString + "<bR>回复人：" & BBSLastPostUser 
	TheString = TheString + "<bR>｜<a  onclick='return cformdelreply();' href='AdminListAction.asp?Code=Reply&ID=" & Rs(5) & "&BBSRoot=" & ID & "&BoardID=" & BBSBoardID & "'>删除帖子</a>｜" 
	
	Response.write TheString
    Count = Count + 1
	rs.movenext
  loop
  
  Response.write "<hr>[<font color=red>" & BBSTitle & "</font>] 合计<font color=red><b>" & RsNo & "</b></font>条  |"
  ThePageUrl	= ThisPage & "?ID=" & ID
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