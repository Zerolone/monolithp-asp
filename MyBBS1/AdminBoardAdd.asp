<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="Inc/Function.asp"-->
<!-- #include file="Inc/Skin.asp"-->
<!-- #include file="TopMenu.asp"-->
<hr>
��ǰҳ�棺�½����
<hr>
<%
  dim ThisPage
  ThisPage = "AdminBoardSetting.asp"
%>
<form action="<%=ThisPage%>" method="post" name="AdminBoardAddFrm" id="AdminBoardAddFrm">
  ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�ƣ� 
  <input name="BoardName" type="text" id="BoardName" size="40">
  <br>
  ��&nbsp;��&nbsp;&nbsp;��&nbsp;�� 
  <select name="BoardRootID" id="BoardRootID">
    <option value="0" selected>��Ϊ����</option>
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
  ˳&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� 
  <input name="BoardOrder" type="text" id="BoardOrder" size="40">
  <br>
  ��&nbsp;��&nbsp;&nbsp;˵&nbsp;���� 
  <textarea name="BoardIntro" cols="38" rows="4" id="BoardIntro">
</textarea>
  <br>
  ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���� 
  <input name="BoardMaster" type="text" id="BoardMaster" size="40">
  <br>
  </p>
  <p>
    <input type="submit" name="Submit" value=" �� �� �� �� ">
  </p>
</form>
<p>&nbsp;</p>
<!-- #include file="BoardTime.asp"-->