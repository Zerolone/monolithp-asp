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
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><table cellspacing="0" width="100%" height="26" cellpadding="0">
	<tr>
		<td align="right" width="100%">
		<a href="ListAddPost.asp?BBS_Board_ID=<%=Board_ID%>&BBS_Root=<%=BBS_ID%>&Board_Name=<%=Board_Name%>">
		<img alt="发表新的主题" src="images/t_reply.gif" border="0" width="89" height="27"></a><a href="ListAdd.asp?Board_ID=<%=Board_ID%>&Board_Name=<%=Board_Name%>"><img alt="发表新的主题" src="images/t_new.gif" border="0"></a></td>
	</tr>
</table>
<%
  Dim ThisPage, FormPage
  ThisPage = "ListSave.asp"
  FormPage	= Request.ServerVariables("HTTP_REFERER")
  
  Dim BBS_Subject

  '-----------------------0---------------------1--------------------------2---------------------------3------------------------4-----------------------5---------------------------6--------------------------7
  Sql= "select [MyBBS_BBS].[BBS_ID], [MyBBS_BBS].[BBS_Subject], [MyBBS_BBS].[BBS_PostUser], [MyBBS_BBS].[BBS_Reply], [MyBBS_BBS].[BBS_View], [MyBBS_BBS].[BBS_PostTime], [MyBBS_BBS].[BBS_Content], [MyBBS_BBS].[BBS_Board_ID], "
	'----------------------------8----------------------------9------------------------------10--------------------------------11
	Sql= Sql & " [Monolith_Users].[Users_ID], [Monolith_Users].[Users_Name], [Monolith_Users].[Users_Subject], [Monolith_Users].[Users_LastLogin] " 
	
	Sql= Sql & " from [MyBBS_BBS], [Monolith_Users] Where [MyBBS_BBS].[BBS_PostUserID] = [Monolith_Users].[Users_ID] And [MyBBS_BBS].[BBS_ID]=" & BBS_ID

  Rs.Open Sql, Conn, 1, 3
  SqlQueryNum	= SqlQueryNum + 1
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

  '-----------------------0---------------------1--------------------------2---------------------------3------------------------4-----------------------5---------------------------6--------------------------7
  Sql= "select [MyBBS_BBS].[BBS_ID], [MyBBS_BBS].[BBS_Subject], [MyBBS_BBS].[BBS_PostUser], [MyBBS_BBS].[BBS_Reply], [MyBBS_BBS].[BBS_View], [MyBBS_BBS].[BBS_PostTime], [MyBBS_BBS].[BBS_Content], [MyBBS_BBS].[BBS_Board_ID], "
	'----------------------------8----------------------------9------------------------------10--------------------------------11
	Sql= Sql & " [Monolith_Users].[Users_ID], [Monolith_Users].[Users_Name], [Monolith_Users].[Users_Subject], [Monolith_Users].[Users_LastLogin] " 
	
	Sql= Sql & " from [MyBBS_BBS], [Monolith_Users] Where [MyBBS_BBS].[BBS_PostUserID] = [Monolith_Users].[Users_ID] And [MyBBS_BBS].[BBS_Root]=" & BBS_ID

  Sql= Sql & " Order By [MyBBS_BBS].[BBS_ID] Asc"
  Rs.open Sql, Conn ,1, 1
  SqlQueryNum	= SqlQueryNum + 1

  Dim i 
  i = 1

  Do while not (rs.bof or rs.eof)
  	TheString =	MyBBS_Template_03_Body
		TheString = Replace(TheString , "{BBS_Subject}",			BBS_Subject)
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
		
		i = i + 1

		Response.write TheString
		Rs.MoveNext
	Loop

  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1

 	TheString =	MyBBS_Template_03_Foot
	TheString = Replace(TheString , "{BBS_ID}", 					BBS_ID)
	TheString = Replace(TheString , "{Board_ID}", 				Board_ID)
	Response.write TheString	

	Call CloseRs
	Call CloseConn

  TheString	= MyBBS_Template_Foot 
  TheString	= Replace(TheString, "{Now}",	Now())
  TheString	= Replace(TheString, "{ExecuteTime}",	FormatNumber((Timer()-Startime)*1000,5) & "毫秒。")
  TheString	= Replace(TheString, "{SqlQueryNum}",	SqlQueryNum)
  Response.write TheString
%>