<!-- #include file="Head.asp"-->
<%
  Dim TheString

  Dim BBS_ID 
  BBS_ID			= Request("BBS_ID")
%>
<!-- #include file="BoardString.asp"-->
<script language="javascript">
function ShowOrHideTable(TableName)
{
	TableName	= eval(TableName);
	if(TableName.style.display == "none")
	{
		TableName.style.display = "";
	}
	else
	{
		TableName.style.display = "none";
	}
}
</script>
<%
  Dim FormPage
  FormPage	= Request.ServerVariables("HTTP_REFERER")
  
  Dim BBS_Subject

  '-----------------------0---------------------1--------------------------2---------------------------3------------------------4-----------------------5---------------------------6--------------------------7
  Sql= "select [MyBBS_BBS].[BBS_ID], [MyBBS_BBS].[BBS_Subject], [MyBBS_BBS].[BBS_PostUser], [MyBBS_BBS].[BBS_Reply], [MyBBS_BBS].[BBS_View], [MyBBS_BBS].[BBS_PostTime], [MyBBS_BBS].[BBS_Content], [MyBBS_BBS].[BBS_Board_ID], "
	'----------------------------8----------------------------9------------------------------10--------------------------------11
	Sql= Sql & " [Monolith_Users].[Users_ID], [Monolith_Users].[Users_Name], [Monolith_Users].[Users_Subject], [Monolith_Users].[Users_LastLogin] " 
	
	Sql= Sql & " from [MyBBS_BBS], [Monolith_Users] Where [MyBBS_BBS].[BBS_PostUserID] = [Monolith_Users].[Users_ID] And [MyBBS_BBS].[BBS_ID]=" & BBS_ID

  Rs.Open Sql, Conn, 1, 3
	if Rs.Bof Or Rs.Eof then
	  Response.write "参数错误！"
	  Response.end
	else
	  Rs(4) = Rs(4) + 1
	  Rs.Update
	  
		TheString = MyBBS_Template_03_Head
		TheString = Replace(TheString , "{BBS_Subject}", 			Rs(1))
		TheString = Replace(TheString , "{BBS_PostUser}",			Rs(2))
		TheString = Replace(TheString , "{BBS_PostUserID}", 	Rs(8))
		TheString = Replace(TheString , "{BBS_PostTime}", 		Rs(5))
		TheString = Replace(TheString , "{BBS_Content}", 			Rs(6))
		TheString = Replace(TheString , "{Users_Subject}", 		Rs(10))
		TheString = Replace(TheString , "{Users_LastLogin}",	Rs(11))
		TheString = Replace(TheString , "{Users_Sign}", 			"用户签名")
		TheString = Replace(TheString , "{BBS_ID}", 					BBS_ID)
		TheString = Replace(TheString , "{Board_ID}", 				Board_ID)
		BBS_Subject	= Rs(1)
		Response.write TheString
		Rs.Close
	end if

  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1


	Dim Page, Pages, j, ThisPage
  ThisPage							= Request.ServerVariables("PATH_INFO")
	Page									= clng(request("Page"))

  '-----------------------0---------------------1--------------------------2---------------------------3------------------------4-----------------------5---------------------------6--------------------------7
  Sql= "select [MyBBS_BBS].[BBS_ID], [MyBBS_BBS].[BBS_Subject], [MyBBS_BBS].[BBS_PostUser], [MyBBS_BBS].[BBS_Reply], [MyBBS_BBS].[BBS_View], [MyBBS_BBS].[BBS_PostTime], [MyBBS_BBS].[BBS_Content], [MyBBS_BBS].[BBS_Board_ID], "
	'----------------------------8----------------------------9------------------------------10--------------------------------11
	Sql= Sql & " [Monolith_Users].[Users_ID], [Monolith_Users].[Users_Name], [Monolith_Users].[Users_Subject], [Monolith_Users].[Users_LastLogin] " 
	Sql= Sql & " from [MyBBS_BBS], [Monolith_Users] Where [MyBBS_BBS].[BBS_PostUserID] = [Monolith_Users].[Users_ID] And [MyBBS_BBS].[BBS_Root]=" & BBS_ID
  Sql= Sql & " Order By [MyBBS_BBS].[BBS_ID] Asc"

  Dim i 
  i = 1
  Dim Page_Break

  Rs.open Sql, Conn, 1, 1
	if Not (Rs.bof Or Rs.Eof) then
		Rs.PageSize=5
		if page=0 then page=1
		pages=rs.pagecount
		if page > pages then page=pages
		rs.AbsolutePage=page
		for j=1 to rs.PageSize 

	  	TheString =	MyBBS_Template_03_Body
			TheString = Replace(TheString , "{BBS_PostUser}",			Rs(2))
			TheString = Replace(TheString , "{BBS_PostUserID}", 	Rs(8))
			TheString = Replace(TheString , "{BBS_PostTime}", 		Rs(5))
			TheString = Replace(TheString , "{BBS_Content}", 			Rs(6))
			TheString = Replace(TheString , "{Users_Subject}", 		Rs(10))
			TheString = Replace(TheString , "{Users_LastLogin}",	Rs(11))
			TheString = Replace(TheString , "{Users_Sign}", 			"用户签名")
			TheString = Replace(TheString , "{BBS_ID}", 					BBS_ID)
			TheString = Replace(TheString , "{Board_ID}", 				Board_ID)
			TheString = Replace(TheString , "{Level}", 						i)
			Response.write TheString
		
			i = i + 1
    	Rs.MoveNext
    	if rs.eof then exit for
  	next

		'Page Break
		Page_Break = "<table  cellspacing='0' width='98%' align='center' height='26' cellpadding='0'>"
		Page_Break = Page_Break & "<form method='Post' action='" & ThisPage & "'>"
		Page_Break = Page_Break & "<tr>"
		Page_Break = Page_Break & "<td align=left>"

	  ThisPage = ThisPage & "?BBS_ID=" & BBS_ID & "&Board_ID=" & Board_ID
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
		Page_Break = Page_Break & "<input type='hidden' name='Board_ID' value='" & Board_ID & "'>"
		Page_Break = Page_Break & "</td>"
		Page_Break = Page_Break & "</tr>"
		Page_Break = Page_Break & "</form>"
		Page_Break = Page_Break & "</table>"
	end if

  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1

 	TheString =	MyBBS_Template_03_Foot
	TheString = Replace(TheString , "{BBS_ID}", 					BBS_ID)
	TheString = Replace(TheString , "{Board_ID}", 				Board_ID)
	Response.write TheString	

	Response.write "<script language='javascript'>"
	Response.write "Page_Break_Head.innerHTML=""" & Page_Break & """" & ";"
	Response.write "Page_Break_Foot.innerHTML=""" & Replace(Page_Break, "left", "right") & """" & ";"
	Response.write "</script>"


	Call CloseRs
	Call CloseConn

  TheString	= MyBBS_Template_Foot 
  TheString	= Replace(TheString, "{Now}",	Now())
  TheString	= Replace(TheString, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "毫秒。")
  TheString	= Replace(TheString, "{SqlQueryNum}",	SqlQueryNum)
  Response.write TheString
%>