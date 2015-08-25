<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="Inc/Function.asp"-->
<!-- #include file="Inc/Skin.asp"-->
<!-- #include file="TopMenu.asp"-->
<hr>
当前页面：关于我们 
<hr>
<%
  dim ThisPage
  
  dim ABoutUs_Person

  ThisPage	= Request.ServerVariables("PATH_INFO")
  if Request("Submit") = " 提 交 修 改 " then 
	ABoutUs_Person	= Trim(Request("ABoutUs_Person"))
  
	Sql = "update [MyBBS_ABoutUs] set [ABoutUS_Person]='" & ABoutUs_Person & "'"
	Conn.execute(Sql)
    Response.write "<script language='javascript'>alert('修改成功！');location.href='" & ThisPage & "';</script>"
  end if

  sql= "select [AboutUS_Person] from [MyBBS_ABoutUs]"
  rs.open sql, conn, 1, 3
  SqlQueryNum	= SqlQueryNum + 1
  if err.number <> 0 then
    response.write "数据库操作失败，请稍候重试！"
	response.end
  else
	if not (rs.bof or rs.eof) then 
	  ABoutUs_Person	= Rs(0)
	end if
  end if
  rs.close
%>
<form action="<%=ThisPage%>" method="post" name="AdminABoutUs" id="AdminABoutUs">
  <textarea name="ABoutUs_Person" cols="100" rows="15" id="ABoutUs_Person"><%=ABoutUs_Person%></textarea>
    <input type="submit" name="Submit" value=" 提 交 修 改 ">
</form>
<!-- #include file="BoardTime.asp"-->