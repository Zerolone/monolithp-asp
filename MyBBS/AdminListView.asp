<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="Inc/Skin.asp"-->
<!-- #include file="TopMenu.asp"-->
<!-- #include file="Inc/AdminJs.asp"-->

<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<hr>
��ǰҳ�棺���Ӳ쿴
<hr>
<%
  dim id 
  id	= Request("ID")
  '--------------0----1-----------2--------------3-----------4----------5------------------6------------------7---------8
  sql= "select [ID], [BBSTitle], [BBSPostUser], [BBSReply], [BBSView], [BBSLastPostTime], [BBSLastPostUser], [BBSBody], [BBSBoardID] from [MyBBS_BBS] Where [ID]=" & ID
'  Response.write Sql
  rs.open sql, conn, 1, 3
  SqlQueryNum	= SqlQueryNum + 1
  if err.number <> 0 then
    response.write "���ݿ����ʧ�ܣ����Ժ����ԣ�"
	response.end
  else
	if (rs.bof or rs.eof) then
	  Response.write "��������"
	  Response.end
    else
      dim BBSTitle, BBSPostUser, BBSReply, BBSView, BBSLastPostTime, BBSLastPostUser, BBSBody, BBSBoardID
	  Rs(4) = Rs(4) + 1
	  ID					= Rs(0)
	  BBSTitle				= Rs(1)
	  BBSPostUser			= Rs(2)
	  BBSReply				= Rs(3)
 	  BBSView				= Rs(4)
	  BBSLastPostTime		= Rs(5)
	  BBSLastPostUser		= Rs(6)
	  BBSBody				= Rs(7)
	  BBSBoardID			= Rs(8)
	  Rs.Update
	  
	  if IsNull(BBSLastPostTime) 		then	BBSLastPostTime		= "��"
	  if IsNull(BBSLastPostUser) 		then	BBSLastPostUser		= "��"
	  
	  TheString = "����ID��" & ID
	  TheString = TheString + "<HR>���⣺" & BBSTitle
	  TheString = TheString + "<br>���ݣ�" & BBSBody
	  TheString = TheString + "<br>�����ˣ�" & BBSPostUser
	  TheString = TheString + "<br>�ظ�����" & BBSReply
	  TheString = TheString + "<br>�������" & BBSView
	  TheString = TheString + "<br>�ظ�ʱ�䣺" & BBSLastPostTime
	  TheString = TheString + "<br>�ظ��ˣ�" & BBSLastPostUser 
	  Response.write TheString
    end if
    rs.close
  end if
  Response.write "<hr>����Ϊ�ظ����ݣ�<br>"
  
  Sql = "Select [BBSTitle], [BBSPostUser], [BBSLastPostTime], [BBSLastPostUser], [BBSBody], [ID] From [MyBBS_BBS] Where [BBSRoot]=" & ID
  Sql = Sql & " And BBSBoardID=" & BBSBoardID
  Sql = Sql & " Order By [ID] Asc"
  Rs.open Sql, Conn ,1, 3

  '-----------��ҳ
  dim page, news, PageSize, RsNo, pgnum, Count, ThePageUrl, ThisPage
  ThisPage	= Request.ServerVariables("PATH_INFO")
  Count = 0
  page=request("page")
  PageSize = 5
  Rs.PageSize = PageSize
  RsNo=Rs.recordcount
  pgnum=Rs.Pagecount
  if page="" or clng(page)<1 then page=1
  if clng(page) > pgnum then page=pgnum
  if pgnum>0 then Rs.AbsolutePage=page						
  '=-=-=--=-=-  

  do while not (rs.bof or rs.eof) and count<Rs.PageSize
	BBSTitle			= Rs(0)
	BBSPostUser			= Rs(1)
	BBSLastPostTime		= Rs(2)
	BBSLastPostUser		= Rs(3)
	BBSBody				= Rs(4)
	
	TheString = "<HR>���⣺" & BBSTitle
	TheString = TheString + "<bR>���ݣ�" & BBSBody
	TheString = TheString + "<bR>�����ˣ�" & BBSPostUser
	TheString = TheString + "<bR>�ظ�ʱ�䣺" & BBSLastPostTime
	TheString = TheString + "<bR>�ظ��ˣ�" & BBSLastPostUser 
	TheString = TheString + "<bR>��<a  onclick='return cformdelreply();' href='AdminListAction.asp?Code=Reply&ID=" & Rs(5) & "&BBSRoot=" & ID & "&BoardID=" & BBSBoardID & "'>ɾ������</a>��" 
	
	Response.write TheString
    Count = Count + 1
	rs.movenext
  loop
  
  Response.write "<hr>[<font color=red>" & BBSTitle & "</font>] �ϼ�<font color=red><b>" & RsNo & "</b></font>��  |"
  ThePageUrl	= ThisPage & "?ID=" & ID
  if page <=1 then
	Response.write "  ��һҳ "
  else
	Response.write "  <a href=""" & ThePageUrl & "&page=" & page-1 & """>��һҳ</a> "
  end if

  if Rs.pagecount-page<1 then
	Response.write "  ��һҳ "
  else
	Response.write "  <a href=""" & ThePageUrl & "&page=" & page+1 & """>��һҳ</a> "
  end if
  Response.write "| ҳ�Σ�<b><font color=red> " & page & "</font>/" & Rs.pagecount & "</b>ҳ <b>" & PageSize & "</b>��/ҳ "
  rs.close

  Response.write "<hr>"
  if Username <> "" then
%>
<form action="ListSave.asp" method="post" name="ReplyTopic">
  <p>�ظ����⣺ 
    <input name="BBSTitle" type="text" size="51">
    <br>
    �ظ����ݣ� 
    <textarea name="BBSBody" cols="50" rows="5"></textarea>
    <input type="hidden" name="BBSBoardID" value="<%=BBSBoardID%>"><input type="hidden" name="BBSRoot" value="<%=ID%>">
    <br>
    <input type="submit" name="Submit" value=" �� �� �� �� �� �� ">
  </p>
</form>
<% end if %>
<!-- #include file="BoardTime.asp"-->