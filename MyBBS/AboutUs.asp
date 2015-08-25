<!-- #include file="Inc/Conn.asp"-->
<!-- #include file="Inc/Config.asp"-->
<!-- #include file="TopMenu.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
  sql= "select [AboutUs_Person] from [MyBBS_AboutUs]"
  rs.open sql, conn, 1, 3
  SqlQueryNum	= SqlQueryNum + 1
  if not (rs.bof or rs.eof) then
    Response.write Rs(0)
  end if
  rs.close
%>
<!-- #include file="BoardTime.asp"-->