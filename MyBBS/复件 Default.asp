<!-- #include file ="../Manage/Inc/Monolith_Conn.asp" -->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="Inc/Skin.asp"-->
<!-- #include file="TopMenu.asp"-->
<span id="MyBBS_Board"></span> 
<%
  dim ID, BoardName, BoardRootID, BoardOrder, BoardIntro, BoardMaster, BoardTopic, BoardReply
  dim BoardLastPostTime, BoardLastPostUser, BoardLastPostTopic, BoardStyle
  dim Arr_BoardStyle
  dim JsString, JsString1
  TheString = ""
  JsString 	= "for(var i=0;i<document.all.length;i++){" + Chr(10)
  JsString1 = ""
  
  '-------------0-----1------------2--------------3-------------4-------------5--------------6-------------7-------------8--------------------9--------------------10--------------------11
  sql= "select [ID], [BoardName], [BoardRootID], [BoardOrder], [BoardIntro], [BoardMaster], [BoardTopic], [BoardReply], [BoardLastPostTime], [BoardLastPostUser], [BoardLastPostTopic], [BoardStyle] from [MyBBS_Board]  order by [BoardRootID], [BoardOrder]"
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
	  BoardStyle			= Rs(11)
  
	  if IsNull(BoardIntro) 		then	BoardIntro 			= "　"
	  if IsNull(BoardMaster) 		then	BoardMaster 		= "　"
	  if IsNull(BoardLastPostTime) 	then	BoardLastPostTime 	= "　"
	  if IsNull(BoardLastPostUser) 	then	BoardLastPostUser 	= "　"
	  if IsNull(BoardLastPostTopic)	then	BoardLastPostTopic 	= "　"

	  if BoardRootID = 0 then
	    Arr_BoardStyle		= Split(BoardStyle,",")
	    '背景图片,Tr颜色一,Tr颜色二,下角颜色
	    '0--------1-------2--------3

		TheString = ""
		TheString = Replace(Skin_BoardMain, "{BoardName}", BoardName)
		TheString = Replace(TheString, "{BoardID}", ID)
		TheString = Replace(TheString, "{StyleBgImage}"	, Trim(Arr_BoardStyle(0)))
		TheString = Replace(TheString, "{StyleTr1}"		, Trim(Arr_BoardStyle(1)))
		TheString = Replace(TheString, "{StyleFoot}"	, Trim(Arr_BoardStyle(3)))
	    Response.write "MyBBS_Board.innerHTML=MyBBS_Board.innerHTML + """ & TheString & """;"+  Chr(10)

		JsString1 = JsString1 + "  var i" & ID & "=0;" + Chr(10)
		JsString = JsString + "if (document.all[i].name=='table" & ID & "'){" + Chr(10)
		JsString = JsString + "  if (i" & ID &" % 2 == 0)"
		JsString = JsString + "  {document.all[i].bgColor='" & Trim(Arr_BoardStyle(2)) & "';}" + Chr(10)
		JsString = JsString + "  else" + Chr(10)
		JsString = JsString + "  {document.all[i].bgColor='" & Trim(Arr_BoardStyle(1)) & "';}" + Chr(10)
		JsString = JsString + "  i" & ID & "= i" & ID  & "+1;" + Chr(10)
		JsString = JsString + "}" + Chr(10)
		


'	  else if Rs(2) <> 0 then
	  else
		TheString = ""
		TheString = Replace(Skin_BoardSecond, "{BoardName}"		,	"<a href='List.asp?BoardID=" & ID & "'>" & BoardName & "</a>")
		TheString = Replace(TheString, "{BoardIntro}"			,	BoardIntro)
		TheString = Replace(TheString, "{BoardMaster}"			,	BoardMaster)
		TheString = Replace(TheString, "{BoardTopic}"			,	BoardTopic)
		TheString = Replace(TheString, "{BoardReply}"			,	BoardReply)
		TheString = Replace(TheString, "{BoardLastPostTime}"	, 	BoardLastPostTime)
		TheString = Replace(TheString, "{BoardLastPostUser}"	, 	BoardLastPostUser)
		TheString = Replace(TheString, "{BoardLastPostTopic}"	,	BoardLastPostTopic)
		TheString = Replace(TheString, "{TableName}"			,	BoardRootID)
		
	    Response.write "MyBBS_Board" & BoardRootID & ".innerHTML=MyBBS_Board" & BoardRootID & ".innerHTML + """ & TheString & """;"+  Chr(10)
	  end if
	  'Response.Flush()
	  rs.movenext
	loop
	JsString = JsString + "}" + Chr(10)

	Response.write JsString1
	Response.write JsString
    Response.write "</script>" & Chr(10)
  end if
  rs.close
%><!-- #include file="BoardTime.asp"-->