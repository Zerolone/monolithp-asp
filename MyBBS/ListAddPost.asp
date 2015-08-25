<!-- #include file="Head.asp"-->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<!-- #include virtual = "/Inc/LoginUserCheck.asp" -->
<%
  if IsLogin = 0 then
    Response.Write "请登陆！"
		Response.end
  end if

  Dim TheString

  Dim ThisPage
  ThisPage = "ListSave.asp"
  
  Dim BBS_Root
  BBS_Root	= Request("BBS_Root")
%>
<!-- #include file="BoardString.asp"-->
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线-->
<%
	Dim BBS_Subject
	'----------------------0
  Sql= "Select Top 1 [BBS_Subject] From [MyBBS_BBS] Where [BBS_ID]=" & BBS_Root
  Rs.Open Sql, Conn, 1, 1
  if Rs.Bof Or Rs.Eof then
  	Response.write "错误的参数！"
  	Response.end
  else
  	BBS_Subject		= Rs(0)
  end if
  Rs.Close
  
  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1

	TheString	= MyBBS_Template_06_Head
	TheString	= Replace(TheString, "{ThisPage}",			ThisPage)
	TheString	= Replace(TheString, "{BBS_Root}",			BBS_Root)
	TheString	= Replace(TheString, "{Board_ID}",			Board_ID)
	TheString	= Replace(TheString, "{BBS_Subject}",		BBS_Subject)
	Response.write TheString

  '-----------------------------------0----------------------------1-------------------------2---------------------------3
  Sql= "Select Top 10 [Monolith_Users].[Users_ID], [Monolith_Users].[Users_Name], [MyBBS_BBS].[BBS_PostTime], [MyBBS_BBS].[BBS_Content] From [MyBBS_BBS], [Monolith_Users] Where [MyBBS_BBS].[BBS_PostUserID] = [Monolith_Users].[Users_ID] And [MyBBS_BBS].[BBS_Root]=" & BBS_Root
  Sql= Sql & " Order By [MyBBS_BBS].[BBS_ID] Asc"
  Rs.open Sql, Conn ,1, 1

  Dim i 
  i = 1

  Do while not (rs.bof or rs.eof)
		TheString	= MyBBS_Template_06_Body
		TheString	= Replace(TheString, "{Users_ID}",			Rs(0))
		TheString	= Replace(TheString, "{Users_Name}",		Rs(1))
		TheString	= Replace(TheString, "{BBS_PostTime}",	Rs(2))
		TheString	= Replace(TheString, "{BBS_Content}",		Rs(3))
		Response.write TheString
		Rs.MoveNext
	Loop
	Rs.Close
	
  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1

	TheString	= MyBBS_Template_06_Foot
	TheString	= Replace(TheString, "{BBS_Root}",			BBS_Root)
	TheString	= Replace(TheString, "{Board_ID}",			Board_ID)
	Response.write TheString

  TheString	= MyBBS_Template_Foot 
  TheString	= Replace(TheString, "{Now}",	Now())
  TheString	= Replace(TheString, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "毫秒。")
  TheString	= Replace(TheString, "{SqlQueryNum}",	SqlQueryNum)
  Response.write TheString
%>