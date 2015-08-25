<!-- #include file="Head.asp"-->
<%
  Dim TheString
%>
<!-- #include file="BoardString.asp"-->
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线-->
<%
  TheString = Replace(MyBBS_Template_02_Head, "{Board_Name}", Board_Name)
  Response.write TheString

  TheString = MyBBS_Template_02_All_Head
  Response.write TheString

	Dim Page_Break
  Dim i, k
  i = 0
  k = 0

	Dim Page, Pages, j, ThisPage
  ThisPage							= Request.ServerVariables("PATH_INFO")
	Page									= clng(request("Page"))

 	'-----------------0-----------1---------------2--------------3--------------4--------------5--------------6-------------7--------------------8-------------9---------------10---------------11--------------12------------------13
	Sql = "Select [BBS_ID], [BBS_Subject], [BBS_Board_ID], [BBS_PostUser], [BBS_PostTime], [BBS_Reply], [BBS_View], [BBS_LastPostUser], [BBS_LastPostTime], [BBS_Icon], [BBS_SubjectDesc], [BBS_IsTop], [BBS_PostUserID], [BBS_LastPostUserID] From [MyBBS_BBS] Where [BBS_Board_ID]=" & Board_ID & " and [BBS_Root]=0 order by [BBS_IsTop] Desc, [BBS_ID] Desc, [BBS_LastPostTime] Desc"

  Rs.open Sql, Conn, 1, 1
	if Not (Rs.bof Or Rs.Eof) then
		Rs.PageSize=10
		if page=0 then page=1
		pages=rs.pagecount
		if page > pages then page=pages
		rs.AbsolutePage=page
		for j=1 to rs.PageSize 
	  	if Rs(11) = 1 then
	  		if i=0 then
	  			TheString = ""
	  			TheString = MyBBS_Template_02_Top_Head
	  			Response.write TheString  		
	  		end if
  		
	  		TheString = MyBBS_Template_02_Top_Body
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
		  
			  Response.write TheString
				i = i + 1
			else
	  		if k=0 then
	  			TheString = ""
	  			TheString = MyBBS_Template_02_Board_Head
	  			Response.write TheString  		
	  		end if
		
	  		TheString = MyBBS_Template_02_Board_Body
	  		TheString = Replace(TheString, "{BBS_ID}", Rs(0))
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
			  	TheString = Replace(TheString, "{BBS_SubjectDesc}",	"<br>" & Rs(10))
			  else
			  	TheString = Replace(TheString, "{BBS_SubjectDesc}",	"")
			  end if
		  
			  Response.write TheString
				k = k + 1
			end if
    	Rs.MoveNext
    	if rs.eof then exit for
  	next
	
		'Page Break
		Page_Break = "<table  cellspacing='0' width='98%' align='center' height='26' cellpadding='0'>"
		Page_Break = Page_Break & "<form method='Post' action='" & ThisPage & "'>"
		Page_Break = Page_Break & "<tr>"
		Page_Break = Page_Break & "<td align=left>"

	  ThisPage = ThisPage & "?Board_ID=" & Board_ID
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
		Page_Break = Page_Break & "共 <b><font color='#FF0000'>" & rs.recordcount & "</font></b> 个主题&nbsp;｜"
		Page_Break = Page_Break & "<b>" & rs.pagesize & "</b> 个主题/页&nbsp;｜"
		Page_Break = Page_Break & "转到：<input type='text' name='page' size=4 maxlength=10 class='InputBox' value='" & page & "'> <input class='InputBox' type='submit'  value=' Goto '  name='cndok'>"
		Page_Break = Page_Break & "<input type='hidden' name='Board_ID' value='" & Board_ID & "'>"
		Page_Break = Page_Break & "</td>"
		Page_Break = Page_Break & "<td align=right><a href='ListAdd.asp?Board_ID=" & Board_ID & "'><img alt='发表新的主题' src='images/t_new.gif' border='0'></a></td>"
		Page_Break = Page_Break & "</tr>"
		Page_Break = Page_Break & "</form>"
		Page_Break = Page_Break & "</table>"
	end if
  
	TheString = ""
  TheString = MyBBS_Template_02_Foot
  Response.write TheString

	Response.write "<script language='javascript'>"
	Response.write "Page_Break_Head.innerHTML=""" & Page_Break & """" & ";"
	Response.write "Page_Break_Foot.innerHTML=""" & Replace(Page_Break, "left", "right") & """" & ";"
	Response.write "</script>"
  
  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1
  
	Call CloseRs
	Call CloseConn

  TheString	= MyBBS_Template_Foot 
  TheString	= Replace(TheString, "{Now}",	Now())
  TheString	= Replace(TheString, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "毫秒。")
  TheString	= Replace(TheString, "{SqlQueryNum}",	SqlQueryNum)
  Response.write TheString
%>