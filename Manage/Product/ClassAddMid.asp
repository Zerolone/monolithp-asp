<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
</head>

<%
	Dim ThisPage, Action , ProductClass_BigName, counter, ProductClass_BigOrder
  ThisPage				= Request.ServerVariables("PATH_INFO")
  Action				= Request("Action")	
  ProductClass_BigName	= Request("ProductClass_BigName")
  
  if Request("Action") = "Add" then
    sql ="Select [ProductClass_MidOrder], [ProductClass_BigOrder], [ProductClass_BigName] from [Monolith_ProductClass] where [ProductClass_BigName]='"&ProductClass_BigName&"' order by ProductClass_MidOrder"
  	Rs.open sql, conn, 1, 3
	  if Not (rs.eof and rs.bof) then
	    Do While Not (Rs.Bof Or Rs.Eof)
	      counter					= rs(0)
	      ProductClass_BigOrder	= rs(1)
	      ProductClass_BigName	= rs(2)
		    Rs.MoveNext
      Loop
	    counter=counter+1
  	else
	    counter=1
  	end if
  	Rs.Close
	
  	Sql = "Select * From [Monolith_ProductClass] Where [ProductClass_BigName]='" & ProductClass_BigName & "' And [ProductClass_MidName]='" & Request.Form("ProductClass_MidName") & "'"
	  Rs.open sql, conn, 1, 3
	  if Not( Rs.Bof or Rs.eof) then
	    Response.write "<script>alert('该类型已经重复！');window.document.location.href='"& ThisPage &"?ProductClass_BigName=" & ProductClass_BigName & "';</script>"
	    Response.end
	  end if
  	Rs.addnew
	  Rs("ProductClass_BigOrder")	= ProductClass_BigOrder
  	Rs("ProductClass_BigName")	= ProductClass_BigName
	  Rs("ProductClass_MidOrder")	= Counter
  	Rs("ProductClass_MidName")	= Request.Form("ProductClass_MidName")
	  Rs.update
	  Call CloseRs
	  Call CloseConn
	  Response.Redirect ThisPage & "?ProductClass_BigName=" & ProductClass_BigName
  end if
%>
<table width="100%" border="0" cellspacing="1" bgcolor="#CCCCCC">
	<form action="<%=ThisPage%>?Action=Add" method="post" name="addmidclass">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;产品管理 &gt;&gt; 产品类别管理 &gt;&gt; 产品中类
			添加</font></b></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right" height="20">
			<font color="#FFFFFF">中类名称：</font></td>
			<td bgcolor="#E3E3E3">
			<input type="text" name="ProductClass_MidName" size="20" maxlength="12" class="InputBox"></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right" height="20">　</td>
			<td bgcolor="#E3E3E3">
			<input type="submit" value=" 提 交 " class="InputBox"></td>
		</tr>
		<input type="hidden" name="ProductClass_BigName" value="<%=ProductClass_BigName%>">
		　</form>
</table>
<%
 	response.write "大类："&ProductClass_BigName&"<br>中类:<br>"
  Sql="Select * from [Monolith_ProductClass] where [ProductClass_BigName]='"&ProductClass_BigName&"' order by [ProductClass_MidOrder] asc"
 	rs.open sql,conn,1,3
 	if not (rs.eof and rs.bof) then
     do while not rs.eof
 		  response.write "&nbsp;&nbsp;"&rs("ProductClass_MidOrder")&"|"&rs("ProductClass_MidName")&"<br>"
  	  rs.movenext
    loop
  end if
 	Call CloseRs
  Call CloseConn
%>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->

</html>
