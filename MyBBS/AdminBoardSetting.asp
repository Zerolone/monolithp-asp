<!-- #include file ="../Manage/Inc/Monolith_Conn.asp" -->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="Inc/Function.asp"-->
<!-- #include file="Inc/Skin.asp"-->
<!-- #include file="TopMenu.asp"-->
<hr>
��ǰҳ�棺�޸İ��
<hr>
<%
  dim ThisPage

  dim ID, BoardName, BoardRootID, BoardOrder, BoardIntro, BoardMaster, BoardTopic, BoardReply
  dim BoardLastPostTime, BoardLastPostUser, BoardLastPostTopic, BoardStyle
  dim Arr_BoardStyle

  ThisPage	= Request.ServerVariables("PATH_INFO")
  if Request("Submit") = " �� �� �� �� �� �� " then 
    ID					= Trim(Request("ID"))
	BoardName			= Trim(Request("BoardName"))
	BoardRootID			= Trim(Request("BoardRootID"))
	BoardOrder			= Trim(Request("BoardOrder"))
	BoardIntro			= Trim(Request("BoardIntro"))
	BoardMaster			= Trim(Request("BoardMaster"))
	BoardTopic			= Trim(Request("BoardTopic"))
	BoardReply			= Trim(Request("BoardReply"))
	BoardLastPostTime	= Trim(Request("BoardLastPostTime"))
	BoardLastPostUser	= Trim(Request("BoardLastPostUser"))
	BoardLastPostTopic	= Trim(Request("BoardLastPostTopic"))
	BoardStyle			= Trim(Request("BoardStyle"))
  
	Sql = "update [MyBBS_Board] set [BoardName]='" & BoardName & "'"
	Sql = Sql & ", [BoardRootID] 		="  & BoardRootID
	Sql = Sql & ", [BoardOrder] 		="  & BoardOrder
	Sql = Sql & ", [BoardIntro] 		='" & BoardIntro & "'"
	Sql = Sql & ", [BoardMaster] 		='" & BoardMaster & "'"
	Sql = Sql & ", [BoardTopic] 		="  & BoardTopic
	Sql = Sql & ", [BoardReply] 		="  & BoardReply
	Sql = Sql & ", [BoardStyle] 		='" & BoardStyle & "'"
	
	if BoardLastPostTime <> "" then Sql = Sql & ", [BoardLastPostTime]	='" & BoardLastPostTime & "'"
	Sql = Sql & ", [BoardLastPostUser]	='" & BoardLastPostUser & "'"
	Sql = Sql & ", [BoardLastPostTopic]	='" & BoardLastPostTopic & "'"
	Sql = Sql & " where [ID]=" & ID
'	Response.write sql & "<hr>"	
'	Response.end
	Conn.execute(Sql)
    Response.write "<script language='javascript'>alert('�޸ĳɹ���');location.href='" & ThisPage & "?ID=" & Trim(ID) & "';</script>"
  end if
  
  if Request("Submit") = " �� �� �� �� " then 
	BoardName			= Trim(Request("BoardName"))
	BoardRootID			= Trim(Request("BoardRootID"))
	BoardOrder			= Trim(Request("BoardOrder"))
	BoardIntro			= Trim(Request("BoardIntro"))
	BoardMaster			= Trim(Request("BoardMaster"))
  
	Sql = "Select * From [MyBBS_Board]"
	rs.open sql, conn, 1, 3
	Rs.AddNew
	Rs("BoardName")		= BoardName
	Rs("BoardRootID")	= BoardRootID
	Rs("BoardOrder")	= BoardOrder
	Rs("BoardIntro")	= BoardIntro
	Rs("BoardMaster")	= BoardMaster
	Rs.Update
	ID = Rs("ID")
	Rs.close
    Response.write "<script language='javascript'>alert('��ӳɹ���');location.href='" & ThisPage & "?ID=" & Trim(ID) & "';</script>"
  end if  

  ID		= Request("ID")
  sql= "select [ID], [BoardName], [BoardRootID], [BoardOrder], [BoardIntro], [BoardMaster], [BoardTopic], [BoardReply], [BoardLastPostTime], [BoardLastPostUser], [BoardLastPostTopic], [BoardStyle] from [MyBBS_Board] Where [ID]=" & ID
  rs.open sql, conn, 1, 3
  SqlQueryNum	= SqlQueryNum + 1
  if err.number <> 0 then
    response.write "���ݿ����ʧ�ܣ����Ժ����ԣ�"
	response.end
  else
	if not (rs.bof or rs.eof) then 
'	  ID					= Rs(0)
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
	end if
  end if
  rs.close
%>
<form action="<%=ThisPage%>" method="post" name="AdminBoardSettingFrm" id="AdminBoardSettingFrm">
  <p>I&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;D�� 
    <input name="ID" type="text" readonly id="ID" value="<%=ID%>" size="40">
    <br>
    ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�ƣ� 
    <input name="BoardName" type="text" id="BoardName" value="<%=BoardName%>" size="40">
    <br>
    ��&nbsp;��&nbsp;&nbsp;��&nbsp;�� 
    <select name="BoardRootID" id="BoardRootID">
      <option value="0" selected>��Ϊ����</option>
      <%
    dim TmpStr
    sql = "Select [ID], [BoardName] From [MyBBS_Board] Where [ID]<> " & ID & " and  [BoardRootID] = 0 Order By [ID] ASC"
	Rs.open Sql, Conn , 1, 3
	SqlQueryNum	= SqlQueryNum + 1
	Do While not (rs.bof or rs.eof)
	  if Rs(0) = BoardRootID then
	    TmpStr = "selected"
	  else
	    TmpStr = ""
	  end if
	  Response.write "<option value='" & Rs(0) & "' " & TmpStr  & ">" & Rs(1) & "</option>"
	  Rs.movenext
	Loop
	Rs.close
  %>
    </select>
    <br>
    ˳&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� 
    <input name="BoardOrder" type="text" id="BoardOrder" value="<%=BoardOrder%>" size="40">
    <br>
    ��&nbsp;��&nbsp;&nbsp;˵&nbsp;���� 
    <textarea name="BoardIntro" cols="38" rows="4" id="BoardIntro"><%=BoardIntro%></textarea>
    <br>
    ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���� 
    <input name="BoardMaster" type="text" id="BoardMaster" value="<%=BoardMaster%>" size="40">
    <br>
    ��&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;���� 
    <input name="BoardTopic" type="text" id="BoardTopic" value="<%=BoardTopic%>" size="40">
    <a href="AdminCount.asp?Code=Topic&ID=<%=ID%>">���¼���������</a> <br>
    ��&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;���� 
    <input name="BoardReply" type="text" id="BoardReply" value="<%=BoardReply%>" size="40">
    <a href="AdminCount.asp?Code=Reply&ID=<%=ID%>">���¼���ظ���</a><br>
    �ϴλظ�ʱ�䣺 
    <input name="BoardLastPostTime" type="text" id="BoardLastPostTime" value="<%=BoardLastPostTime%>" size="40">
    <br>
    �ϴλظ����ߣ� 
    <input name="BoardLastPostUser" type="text" id="BoardLastPostUser" value="<%=BoardLastPostUser%>" size="40">
    <br>
    �ϴλظ����⣺ 
    <input name="BoardLastPostTopic" type="text" id="BoardLastPostTopic" value="<%=BoardLastPostTopic%>" size="40">
    <br>
	<%
'	  if BoardRootID = 0 then 
	    Arr_BoardStyle		= Split(BoardStyle, ",")
	%>
    ��&nbsp;��&nbsp;&nbsp;ͼ&nbsp;Ƭ�� 
    <input name="BoardStyle" type="text" id="BoardStyle" value="<%=Trim(Arr_BoardStyle(0))%>" size="40">
    <br>
    ��&nbsp;��&nbsp;&nbsp;ɫ&nbsp;һ�� 
    <input name="BoardStyle" type="text" id="BoardStyle" value="<%=Trim(Arr_BoardStyle(1))%>" size="40">
    <br>
    ��&nbsp;��&nbsp;&nbsp;ɫ&nbsp;���� 
    <input name="BoardStyle" type="text" id="BoardStyle" value="<%=Trim(Arr_BoardStyle(2))%>" size="40">
    <br>
    ��&nbsp;��&nbsp;&nbsp;��&nbsp;ɫ�� 
    <input name="BoardStyle" type="text" id="BoardStyle" value="<%=Trim(Arr_BoardStyle(3))%>" size="40">
    <br>
	<% 'end if %>
  </p>
    </p>
  <p>
    <input type="submit" name="Submit" value=" �� �� �� �� �� �� ">
  </p>
</form>
<p>��</p>
<!-- #include file="BoardTime.asp"-->