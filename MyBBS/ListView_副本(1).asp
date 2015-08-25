<!-- #include file="Head.asp"-->
<%
  Dim Board_ID, Board_Name
  Board_ID 		= Request("Board_ID")
  Board_Name	= Request("Board_Name")
  
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
%>

<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><table width="100%" cellspacing="0" cellpadding="0" style="border: 1px solid #000000">
	<tr>
		<td background="images/tile_back.gif" height="32">
		<img height="8" src="images/nav_m.gif" width="8" border="0"> <b>
		<font color="#FFFFFF"><%=Rs(1)%></font></b> </td>
	</tr>
	<tr>
		<td align="right" width="100%" background="images/tile_sub.gif" height="24">
		<a href="#">跟踪主题</a> | <a href="#">邮寄主题</a> | <a href="#">打印主题</a>&nbsp;
		</td>
	</tr>
	<tr>
		<td>
		<table cellspacing="1" cellpadding="4" width="100%" border="0">
			<tr>
				<td class="pformstrip" width="21%" height="24">
				<a href="ViewUser.asp?Users_ID=<%=Rs(8)%>"><%=Rs(9)%></a></td>
				<td class="pformstrip" width="79%" height="24">
				<img src="../images/to_post_off.gif" border="0"> <%=Rs(5)%> </td>
			</tr>
			<tr>
				<td class="row4" width="1%" valign="top"><br>
				发表总数: <%=Rs(10)%><br>
				最后登陆: <%=Rs(11)%><br>
				会员编号: <%=Rs(8)%><br>
				　</span></td>
				<td width="99%" class="row4" valign="top"><%=Rs(6)%><br>
				<br>
				--------------------<br>
				签名。。。 </td>
			</tr>
			<tr>
				<td class="pformstrip" width="21%" align="center">
				　</td>
				<td class="pformstrip" width="79%">
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td><a href="#">
						<img alt="Go to the top of the page" src="../images/p_up.gif" border="0"></a></td>
						<td align="right">
						<a title="快速回复这个帖子" href="ListAddPost.asp?BBS_Board_ID=<%=Board_ID%>&BBS_Root=<%=BBS_ID%>&Board_Name=<%=Board_Name%>"><img alt="Quote Post" src="../images/p_quote.gif" border="0"></a></td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#8394B2" height="5">
		<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--></td>
	</tr>
	<%
		Rs.Close
	end if

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
%>
	<tr>
		<td>
		<table cellspacing="1" cellpadding="4" width="100%" border="0">
			<tr>
				<td class="pformstrip" width="21%" height="24">
				<a href="ViewUser.asp?Users_ID=<%=Rs(8)%>"><%=Rs(9)%></a></td>
				<td class="pformstrip" width="79%" height="24">
				<table border="0" width="100%">
					<tr>
						<td><img src="../images/to_post_off.gif" border="0"> <%=Rs(5)%></td>
						<td align="right">&nbsp;发表于:#<%=i%></td>
					</tr>
				</table>
				</td>
			</tr>
			<tr>
				<td class="row4" width="1%" valign="top"><br>
				发表总数: <%=Rs(10)%><br>
				最后登陆: <%=Rs(11)%><br>
				会员编号: <%=Rs(8)%><br>
				　</span></td>
				<td width="99%" class="row4" valign="top"><%=Rs(6)%><br>
				<br>
				--------------------<br>
				签名。。。 </td>
			</tr>
			<tr>
				<td class="pformstrip" width="21%" align="center">
				　</td>
				<td class="pformstrip" width="79%">
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td><a href="#">
						<img alt="Go to the top of the page" src="../images/p_up.gif" border="0"></a></td>
						<td align="right">
						<a title="快速回复这个帖子" href="ListAddPost.asp?BBS_Board_ID=<%=Board_ID%>&BBS_Root=<%=BBS_ID%>&Board_Name=<%=Board_Name%>"><img alt="Quote Post" src="../images/p_quote.gif" border="0"></a></td>
					</tr>
				</table>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#8394B2" height="5">
		<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--></td>
	</tr>
	<%
    Rs.Movenext
    i = i + 1
  Loop
%>
</table>
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><table cellspacing="0" width="100%" height="26" cellpadding="0">
	<tr>
		<td align="right" width="100%">
		<a title="开启快速回复窗口" href="javascript:ShowOrHideTable('FrmFastReply');">
		<img alt="Fast Reply" src="images/t_qr.gif" border="0" width="89" height="27"></a><a href="ListAddPost.asp?BBS_Board_ID=<%=Board_ID%>&BBS_Root=<%=BBS_ID%>&Board_Name=<%=Board_Name%>"><img alt="发表新的主题" src="images/t_reply.gif" border="0" width="89" height="27"></a><a href="ListAdd.asp?Board_ID=<%=Board_ID%>&Board_Name=<%=Board_Name%>"><img alt="发表新的主题" src="images/t_new.gif" border="0"></a></td>
	</tr>
</table>
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><table id="FrmFastReply" width="100%" cellspacing="0" cellpadding="0" style="border: 1px solid #000000; display:none;">
	<form action="ListSave.asp" method="post" name="ReplyTopic">
		<input name="BBS_Board_ID" type="hidden" value="<%=Board_ID%>">
		<input name="BBS_Root" type="hidden" value="<%=BBS_ID%>">
		<tr>
			<td background="images/tile_back.gif" height="32">
			<img height="8" src="images/nav_m.gif" width="8" border="0"> <b>
			<font color="#FFFFFF">快速回复</font></b> </td>
		</tr>
		<tr>
			<td align="center">
			<!--分割线-->
			<img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><br>
			<textarea class="InputBox" tabindex="1" name="BBS_Content" rows="8" cols="70"></textarea><br>
			<br>
			<input class="InputBox" type="Checkbox" checked value="yes" name="enableemo"> 
			启用表情图标 |
			<input class="InputBox" type="CheckBox" checked value="yes" name="enablesig"> 
			启用签名档<br>
			<br>
			<input class="InputBox" accesskey="s" tabindex="2" type="submit" value=" 加 入 回 复 " name="submit">&nbsp;
			<input class="InputBox" type="button" value="更多选项" name="preview">&nbsp;&nbsp;&nbsp;
			<input class="InputBox" onclick="ShowOrHideTable('FrmFastReply');" type="button" value="关闭快速回复" name="qrc"><br>
			<!--分割线-->
			<img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线-->
			</td>
		</tr>
	</form>
</table>
<!-- #include file="CopyRight.asp"-->