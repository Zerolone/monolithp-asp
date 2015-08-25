<!-- #include file ="../Inc/Monolith_Conn.asp" -->
<!-- #include file="Head.asp"-->
<!-- #include file="Skin.asp"--><%
  Dim Board_ID, Board_Name
  Board_ID 		= Request("BBS_Board_ID")
  Board_Name	= Request("Board_Name")
  
  Dim TheString

  Dim ThisPage, FormPage
  ThisPage = "ListSave.asp"
  FormPage	= Request.ServerVariables("HTTP_REFERER")
  
  Dim BBS_Root
  BBS_Root	= Request("BBS_Root")
%>
<!-- #include file="BoardString.asp"-->
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><table width="100%" cellspacing="0" cellpadding="0" style="border: 1px solid #000000">
	<form action="<%=ThisPage%>" method="post" name="ListAddFrm" id="ListAddFrm">
		<input name="BBS_Board_ID" type="hidden" value="<%=Board_ID%>">
		<input name="BBS_Root" type="hidden" value="<%=BBS_Root%>">
		<input name="FormPage" type="hidden" value="<%=FormPage%>">
		<tr>
			<td background="images/tile_back.gif" height="32" colspan="2"><b>
			<font color="#FFFFFF">&nbsp;&nbsp; 回复 <%=BBS_Subject%> 主题</font></b>
			</td>
		</tr>
		<tr>
			<td class="pformstrip" colspan="2">主题内容</td>
		</tr>
		<tr>
			<td class="pformleft" align="middle" width="23%">　</td>
			<td class="pformright" valign="top" width="77%">
			<textarea class="InputBox" tabindex="5" name="BBS_Content" rows="20" cols="80"></textarea> 
			*</td>
		</tr>
		<tr>
			<td class="pformleft" width="23%"><b>主题图标</b></td>
			<td class="pformright" width="77%">
			<input type="radio" value="1" name="BBS_Icon" checked class="InputBox">&nbsp;
			<img src="folder_post_icons/icon1.gif" align="middle">&nbsp;&nbsp;
			<input type="radio" value="2" name="BBS_Icon" class="InputBox">&nbsp;
			<img src="folder_post_icons/icon2.gif" align="middle">&nbsp;&nbsp;
			<input type="radio" value="3" name="BBS_Icon" class="InputBox">&nbsp;
			<img src="folder_post_icons/icon3.gif" align="middle">&nbsp;&nbsp;
			<input type="radio" value="4" name="BBS_Icon" class="InputBox">&nbsp;
			<img src="folder_post_icons/icon4.gif" align="middle">&nbsp;&nbsp;
			<input type="radio" value="5" name="BBS_Icon" class="InputBox">&nbsp;
			<img src="folder_post_icons/icon5.gif" align="middle">&nbsp;&nbsp;
			<input type="radio" value="6" name="BBS_Icon" class="InputBox">&nbsp;
			<img src="folder_post_icons/icon6.gif" align="middle">&nbsp;&nbsp;
			<input type="radio" value="7" name="BBS_Icon" class="InputBox">&nbsp;
			<img src="folder_post_icons/icon7.gif" align="middle"><br>
			<input type="radio" value="8" name="BBS_Icon" class="InputBox">&nbsp;
			<img src="folder_post_icons/icon8.gif" align="middle">&nbsp;&nbsp;
			<input type="radio" value="9" name="BBS_Icon" class="InputBox">&nbsp;
			<img src="folder_post_icons/icon9.gif" align="middle">&nbsp;&nbsp;
			<input type="radio" value="10" name="BBS_Icon" class="InputBox">&nbsp;
			<img src="folder_post_icons/icon10.gif" align="middle">&nbsp;&nbsp;
			<input type="radio" value="11" name="BBS_Icon" class="InputBox">&nbsp;
			<img src="folder_post_icons/icon11.gif" align="middle">&nbsp;&nbsp;
			<input type="radio" value="12" name="BBS_Icon" class="InputBox">&nbsp;
			<img src="folder_post_icons/icon12.gif" align="middle">&nbsp;&nbsp;
			<input type="radio" value="13" name="BBS_Icon" class="InputBox">&nbsp;
			<img src="folder_post_icons/icon13.gif" align="middle">&nbsp;&nbsp;
			<input type="radio" value="14" name="BBS_Icon" class="InputBox">&nbsp;
			<img src="folder_post_icons/icon14.gif" align="middle"><br>
			<input type="radio" value="0" name="BBS_Icon" class="InputBox">&nbsp; 
			[ Use None ] </td>
		</tr>
		<tr>
			<td class="pformstrip" colspan="2">
			<p align="center">
			<input class="InputBox" tabindex="7" type="submit" value=" 加 入 回 复 " name="submit">&nbsp;
			<input class="InputBox" tabindex="8" type="submit" value=" 预 览 主 题 " name="preview">
			</p>
			</td>
		</tr>
	</form>
</table>
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><table width="100%" cellspacing="0" cellpadding="5" style="border: 1px solid #000000">
	<tr>
		<td background="images/tile_back.gif" height="32" colspan="2"><b>
		<font color="#ffffff">&nbsp;&nbsp; </font><font color="#FFFFFF">最近十个回帖 [ 
		按时间顺序 ]</font></b> </td>
	</tr>
<%
  '-----------------------0---------------------1--------------------------2---------------------------3------------------------4-----------------------5---------------------------6--------------------------7
  Sql= "Select Top 10 [MyBBS_BBS].[BBS_ID], [MyBBS_BBS].[BBS_Subject], [MyBBS_BBS].[BBS_PostUser], [MyBBS_BBS].[BBS_Reply], [MyBBS_BBS].[BBS_View], [MyBBS_BBS].[BBS_PostTime], [MyBBS_BBS].[BBS_Content], [MyBBS_BBS].[BBS_Board_ID], "
	'----------------------------8----------------------------9------------------------------10--------------------------------11
	Sql= Sql & " [Monolith_Users].[Users_ID], [Monolith_Users].[Users_Name], [Monolith_Users].[Users_Subject], [Monolith_Users].[Users_LastLogin] " 
	
	Sql= Sql & " from [MyBBS_BBS], [Monolith_Users] Where [MyBBS_BBS].[BBS_PostUserID] = [Monolith_Users].[Users_ID] And [MyBBS_BBS].[BBS_Root]=" & BBS_Root

  Sql= Sql & " Order By [MyBBS_BBS].[BBS_ID] Asc"
  Rs.open Sql, Conn ,1, 1
  SqlQueryNum	= SqlQueryNum + 1

  Dim i 
  i = 1

  Do while not (rs.bof or rs.eof)
%>
	<tr>
		<td class="pformstrip" width="21%" height="24">
		<a href="ViewUser.asp?Users_ID=<%=Rs(8)%>"><%=Rs(9)%></a></td>
		<td class="pformstrip" width="79%" height="24">
		<img src="../images/to_post_off.gif" border="0"> <%=Rs(5)%> </td>
	</tr>
	<tr>
		<td class="row4" width="1%" valign="top">　</td>
		<td width="99%" class="row4" valign="top"><%=Rs(6)%></td>
	</tr>
		<%
    Rs.Movenext
    i = i + 1
  Loop
%>
	<tr>
		<td class="pformstrip" width="100%" align="center" colspan="2">
		<a href="#">察看所有回复(在新窗口中)</a></td>
	</tr>
	<tr>
		<td bgcolor="#8394B2" height="5" colspan="2">
		<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--></td>
	</tr>
</table>
<!-- #include file="CopyRight.asp"-->