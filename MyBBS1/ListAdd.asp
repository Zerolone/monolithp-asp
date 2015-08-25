<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="Inc/Function.asp"-->
<!-- #include file="Inc/Skin.asp"-->
<!-- #include file="TopMenu.asp"-->
<meta http-equiv=Content-Type content="text/html; charset=gb2312">

<hr>
当前页面：发表帖子
<hr>
<%
  if Username="" then
    Response.write "请先登录或者注册！"
	Response.end
  end if

  dim BoardID
   
  BoardID		= Request("BoardID")

  dim ThisPage, FormPage
  ThisPage = "ListSave.asp"
  FormPage	= Request.ServerVariables("HTTP_REFERER")

%>
<form action="<%=ThisPage%>" method="post" name="ListAddFrm" id="ListAddFrm">
  <input name="BBSBoardID" type="hidden" value="<%=BoardID%>">
  <input name="FormPage" type="hidden" value="<%=FormPage%>">

  标题： 
  <input name="BBSTitle" type="text" id="BBSTitle" size="40">
  <br>
  内容 ： 
  <textarea name="BBSBody" cols="38" rows="4" id="BBSBody">
</textarea></p>
  <p>
    <input type="submit" name="Submit" value=" 添 加 帖 子 ">
  </p>
</form>
<p>&nbsp;</p>
<!-- #include file="BoardTime.asp"-->