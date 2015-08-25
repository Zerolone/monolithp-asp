<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
</head>
<form action="<%=ThisPage%>?Action=Add" method="post" name="addmidclass">
	<table width="100%" border="0" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2">
			<b><font color="#FFFFFF">&nbsp;产品管理 &gt;&gt; 产品类别管理 &gt;&gt; 产品大类修改</font></b></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right" height="20">
			<font color="#FFFFFF">大类名称：</font></td>
			<td bgcolor="#E3E3E3">
			<input type="text" name="ProductClass_BigName" size="20" maxlength="12" class="InputBox"></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right" height="20">
			　</td>
			<td bgcolor="#E3E3E3">
			<input type="submit" value=" 提 交 " class="InputBox"></td>
		</tr>
	</table>
</form>
<%
	Dim ThisPage, Action, counter
  ThisPage	= Request.ServerVariables("PATH_INFO")
  Action	= Request("Action")		
  if Request("Action") = "Add" then
    sql="Select Max(ProductClass_BigOrder) from [Monolith_ProductClass]"
    Rs.Open Sql, Conn, 1, 3
    if Not (rs.eof and rs.bof) then
	  counter=Rs(0)+1
    else
	  counter=1
    end if
    Rs.Close

    sql="Select * FROM [Monolith_ProductClass] Where [ProductClass_BigName] ='" & Request.Form("ProductClass_BigName") & "'"
	Rs.open sql, conn, 1, 3
	if Not( Rs.Bof or Rs.eof) then
	  Response.write "<script>alert('该类型已经重复！');window.document.location.href='"& ThisPage & "';</script>"
	  Resposne.end
	end if
	Rs.addnew
	Rs("ProductClass_BigOrder")	= counter
	Rs("ProductClass_MidOrder")	= 1
	Rs("ProductClass_BigName")	= Request.Form("ProductClass_BigName")
	Rs("ProductClass_MidName")	= "临时分类"
	Rs.update
	Rs.close
	set Rs=nothing
	Response.write "<script>window.document.location.href='"& ThisPage & "';</script>"
  else
	response.write "<br>&nbsp;现有大类："
	sql="Select Distinct [ProductClass_BigName],[ProductClass_BigOrder] from [Monolith_ProductClass] order by [ProductClass_BigOrder] Asc"
	rs.open sql,conn,1,1
	if not (rs.eof and rs.bof) then
      do while not rs.eof
	    response.write "&nbsp;&nbsp;"&rs("ProductClass_BigName")
	    rs.movenext
	  loop
	end if
	Call CloseRs
	Call CloseConn
  end if
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->