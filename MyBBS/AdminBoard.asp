<!-- #include file ="../Manage/Inc/Monolith_Conn.asp" -->
<!-- #include file="Inc/Function.asp"-->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="Inc/Skin.asp"-->
<!-- #include file="TopMenu.asp"-->
<!-- #include file="Inc/AdminJs.asp"-->

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<hr>
当前页面：修改板块
<hr>
<form name="AdminBoardFrm" action="AdminBoardSave.asp" method="post">
  <input name="Submit" type="submit" id="Submit" value=" 修 改 内 容 ">
    <span id="MyBBS_Board"></span> 
  <input name="Submit" type="submit" id="Submit" value=" 修 改 内 容 ">
</form>
<%
  dim ID, BoardName, BoardRootID, BoardOrder, BoardIntro, BoardMaster, BoardTopic, BoardReply, BoardLastPostTime, BoardLastPostUser, BoardLastPostTopic
  TheString = ""
  '-------------0-----1------------2--------------3-------------4-------------5--------------6-------------7-------------8--------------------9--------------------10-------------------
  sql= "select [ID], [BoardName], [BoardRootID], [BoardOrder], [BoardIntro], [BoardMaster], [BoardTopic], [BoardReply], [BoardLastPostTime], [BoardLastPostUser], [BoardLastPostTopic] from [MyBBS_Board]  order by [BoardRootID], [BoardOrder]"
  rs.open sql, conn, 1, 3
  SqlQueryNum	= SqlQueryNum + 1
  if err.number <> 0 then
    response.write "数据库操作失败，请稍候重试！"
	response.end
  else
    Response.write "<script language=javascript>" & Chr(10)
	do while not (rs.bof or rs.eof)
	  ID					= Rs(0)
	  BoardName				= Rs(1)
	  BoardRootID			= Rs(2)
	  BoardOrder			= Rs(3)
	  BoardIntro			= Rs(4)
	  BoardMaster			= Rs(5)
	  BoardTopic			= Rs(6)
	  BoardReply			= Rs(7)
	  BoardLastPostTime		= Rs(8)
	  BoardLastPostUser		= Rs(9)
	  BoardLastPostTopic	= Rs(10)
  
	  if BoardRootID = 0 then
		TheString = ""
		TheString = Replace(Skin_AdminBoardMain, "{BoardName}", "<input name='BoardName' type='text' size='20' value='" & BoardName & "'>")
		TheString = Replace(TheString, "{BoardOrder}"		,	"<input name='ID' type='hidden' value='" & ID & "'><input name='BoardOrder' type='text' size='4' value='" & BoardOrder & "'>")
		TheString = Replace(TheString, "{BoardTopic}"		,	"<input name='BoardTopic' 	type='text' 	size='4' 	value='" & BoardTopic & "'>")
		TheString = Replace(TheString, "{BoardSetting}"		,	"<a href='AdminBoardSetting.asp?ID=" & ID & "'>修改设置")
		TheString = Replace(TheString, "{BoardDel}"					,	"<a  onclick='return cform();' href='#'>删除板块（<font color=red>功能制作中</font>）")
		TheString = Replace(TheString, "{BoardID}", ID)
	    Response.write "MyBBS_Board.innerHTML=MyBBS_Board.innerHTML + """ & TheString & """;"+  Chr(10)
	  else
		TheString = ""
		TheString = Replace(Skin_AdminBoardSecond, "{BoardName}"	,	"<input name='BoardName'	type='text'		size='20' 	value='" & BoardName & "'>")
		TheString = Replace(TheString, "{BoardOrder}"				,	"<input name='ID' 			type='hidden' 				value='" & ID & "'><input name='BoardOrder' type='text' size='4' value='" & BoardOrder & "'>")
		TheString = Replace(TheString, "{BoardSetting}"				,	"<a href='AdminBoardSetting.asp?ID=" & ID & "'>修改设置")
		TheString = Replace(TheString, "{BoardDel}"					,	"<a href='AdminList.asp?BoardID=" & ID &"'>察看帖子</a>｜<a  onclick='return cform();' href='#'>删除板块（<font color=red>功能制作中</font>）")
		'TheString = Replace(TheString, "{BoardLastPostTime}"	, 	BoardLastPostTime)
		'TheString = Replace(TheString, "{BoardLastPostUser}"	, 	BoardLastPostUser)
		'TheString = Replace(TheString, "{BoardLastPostTopic}"	,	BoardLastPostTopic)
	    Response.write "MyBBS_Board" & BoardRootID & ".innerHTML=MyBBS_Board" & BoardRootID & ".innerHTML + """ & TheString & """;"+  Chr(10)
	  end if
	  rs.movenext
	loop
    Response.write "</script>" & Chr(10)
  end if
  rs.close
%><!-- #include file="BoardTime.asp"-->