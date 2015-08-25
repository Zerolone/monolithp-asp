<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="Inc/Function.asp"-->
<!-- #include file="Inc/Skin.asp"-->
<!-- #include file="TopMenu.asp"-->
<hr>
当前页面：新建板块
<hr>
<%
  dim ThisPage
  ThisPage = "AdminBoardSetting.asp"
%>
<form action="<%=ThisPage%>" method="post" name="AdminBoardAddFrm" id="AdminBoardAddFrm">
  名&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;称： 
  <input name="BoardName" type="text" id="BoardName" size="40">
  <br>
  所&nbsp;属&nbsp;&nbsp;类&nbsp;别： 
  <select name="BoardRootID" id="BoardRootID">
    <option value="0" selected>作为分类</option>
  <%
    dim TmpStr
    sql = "Select [ID], [BoardName] From [MyBBS_Board] Where [BoardRootID] = 0 Order By [ID] ASC"
	Rs.open Sql, Conn , 1, 3
	SqlQueryNum	= SqlQueryNum + 1
	Do While not (rs.bof or rs.eof)
	  Response.write "<option value='" & Rs(0) & "'>" & Rs(1) & "</option>"
	  Rs.movenext
	Loop
	Rs.close
  %>
  </select>
  <br>
  顺&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;序： 
  <input name="BoardOrder" type="text" id="BoardOrder" size="40">
  <br>
  板&nbsp;块&nbsp;&nbsp;说&nbsp;明： 
  <textarea name="BoardIntro" cols="38" rows="4" id="BoardIntro">
</textarea>
  <br>
  版&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;主： 
  <input name="BoardMaster" type="text" id="BoardMaster" size="40">
  <br>
  </p>
  <p>
    <input type="submit" name="Submit" value=" 添 加 板 块 ">
  </p>
</form>
<p>&nbsp;</p>
<!-- #include file="BoardTime.asp"-->