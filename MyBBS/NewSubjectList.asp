<!-- #include virtual="/MyBBS/Head.asp"-->
<%
  Dim TheString
%>
<!-- #include virtual="/MyBBS/BoardString.asp"-->
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><%
  TheString = MyBBS_Template_05_Head
  Response.write TheString


  TheString = ""
 	'------------------------0-----------1---------------2--------------3--------------4--------------5--------------6-------------7--------------------8-------------9---------------10---------------11--------------12------------------13
	Sql = "Select Top 20 [BBS_ID], [BBS_Subject], [BBS_Board_ID], [BBS_PostUser], [BBS_PostTime], [BBS_Reply], [BBS_View], [BBS_LastPostUser], [BBS_LastPostTime], [BBS_Icon], [BBS_SubjectDesc], [BBS_IsTop], [BBS_PostUserID], [BBS_LastPostUserID] From [MyBBS_BBS] Where [BBS_Root]=0 order by [BBS_ID] Desc, [BBS_LastPostTime] Desc"

  Rs.open Sql, Conn, 1, 1
  SqlQueryNum	= SqlQueryNum + 1
  
  Do While Not (Rs.Bof Or Rs.Eof)
 		TheString = MyBBS_Template_05_Body
		TheString = Replace(TheString, "{BBS_ID}", 						Rs(0))
	  TheString = Replace(TheString, "{BBS_Subject}"	,			Rs(1))
	  TheString = Replace(TheString, "{Board_ID}"	,					Rs(2))
	  TheString = Replace(TheString, "{BBS_PostUser}"	,			Rs(3))
	  TheString = Replace(TheString, "{BBS_PostTime}"	,			Rs(4))
	  TheString = Replace(TheString, "{BBS_Reply}"		,			Rs(5))
	  TheString = Replace(TheString, "{BBS_View}"			,			Rs(6))
	  TheString = Replace(TheString, "{BBS_LastPostUser}",	Rs(7))
	  TheString = Replace(TheString, "{BBS_LastPostTime}",	Rs(8))
	  TheString = Replace(TheString, "{BBS_PostUserID}",		Rs(12))
	  TheString = Replace(TheString, "{BBS_LastPostUserID}",Rs(13))
	  
	  if Rs(9) > 0  then
	  	TheString = Replace(TheString, "{BBS_Icon}",	Rs(9))
	  else
	  	TheString = Replace(TheString, "{BBS_Icon}",	"0")
	  end if
 
	  if Rs(10) <> ""  then
	  	TheString = Replace(TheString, "{BBS_SubjectDesc}", "<br>" & Rs(10))
	  else
	  	TheString = Replace(TheString, "{BBS_SubjectDesc}",	"")
	  end if
		  
	  TheString = Replace(TheString, "{Board_Name}",	Board_Name)

	  Response.write TheString
	  Rs.MoveNext
	loop
	
	TheString = ""
  TheString = MyBBS_Template_05_Foot
  Response.write TheString
  
	Call CloseRs
	Call CloseConn
%><!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><%
  TheString	= MyBBS_Template_Foot 
  TheString	= Replace(TheString, "{Now}",	Now())
  TheString	= Replace(TheString, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "毫秒。")
  TheString	= Replace(TheString, "{SqlQueryNum}",	SqlQueryNum)
  Response.write TheString
%>