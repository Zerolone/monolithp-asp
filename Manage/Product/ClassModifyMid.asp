<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
</head>

<%
	Dim ThisPage, Action , ProductClass_MidName, New_ProductClass_MidName, Old_ProductClass_MidName, ProductClass_BigName


  ThisPage					= Request.ServerVariables("PATH_INFO")
  Action					= trim(Request("Action"))
  ProductClass_BigName		= trim(Request("ProductClass_BigName"))
  ProductClass_MidName		= trim(Request("ProductClass_MidName"))
  New_ProductClass_MidName	= trim(Request.form("New_ProductClass_MidName"))
  Old_ProductClass_MidName	= trim(Request.form("Old_ProductClass_MidName"))
  if ProductClass_BigName ="" or ProductClass_MidName = ""  then Response.Redirect("Admin_Product_ClassManage.asp")
  Sql = "Select * from [Monolith_ProductClass] where [ProductClass_BigName] = '" & ProductClass_BigName & "' And [ProductClass_MidName]='" & ProductClass_MidName & "'"
  rs.Open Sql, conn, 1, 3
  if rs.Bof or rs.EOF then
	Response.write "<br><li>此小类不存在！</li>"
  else
	if Action="Modify" then
      rs("ProductClass_MidName") = New_ProductClass_MidName
      rs.update
	  rs.Close
	  Set rs=Nothing
      if New_ProductClass_MidName <> Old_ProductClass_MidName then
        Sql = "Update [Monolith_ProductClass] set [ProductClass_MidName]='" & New_ProductClass_MidName & "' where [ProductClass_BigName]='" & ProductClass_BigName & "' and [ProductClass_MidName]='" & Old_ProductClass_MidName & "'"
		conn.execute sql
        Sql = "Update [Monolith_Product] set [Product_MidName]='" & New_ProductClass_MidName & "' where [Product_BigName]='" & ProductClass_BigName & "' and [Product_MidName]='" & Old_ProductClass_MidName & "'"
		conn.execute sql
	  end if

	Response.write "<script language=""javascript"">"
	Response.write "  alert(""小类修改成功，请刷新页面。"");"
	Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
	Response.write "</script>"

	end if
  end if
%>
<table width="100%" border="0" cellspacing="1" bgcolor="#CCCCCC">
	<form name="form1" method="post" action="<%=ThisPage%>?Action=Modify">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;产品管理 &gt;&gt; 产品类别管理 &gt;&gt; 产品中类修改</font></b></td>
		</tr>
		<tr>
			<td width="112" height="22" bgcolor="#999999" align="right">
			<font color="#FFFFFF"><strong>所属大类：</strong></font></td>
			<td width="836" bgcolor="#FFFFFF"><%=rs("ProductClass_BigName")%>
			<input name="Old_ProductClass_MidName" type="hidden" id="Old_ProductClass_MidName" value="<%=rs("ProductClass_MidName")%>">
			<input name="ProductClass_BigName" type="hidden" id="ProductClass_BigName" value="<%=rs("ProductClass_BigName")%>">
			<input name="ProductClass_MidName" type="hidden" id="ProductClass_MidName" value="<%=rs("ProductClass_MidName")%>">
			</td>
		</tr>
		<tr>
			<td height="22" bgcolor="#999999" align="right">
			<font color="#FFFFFF"><strong>小类名称：</strong></font></td>
			<td bgcolor="#FFFFFF">
			<input name="New_ProductClass_MidName" type="text" id="New_ProductClass_MidName" value="<%=rs("ProductClass_MidName")%>" size="20" maxlength="30" class="InputBox"></td>
		</tr>
		<tr>
			<td height="22" align="center" bgcolor="#999999">　</td>
			<td align="center" bgcolor="#FFFFFF">
			<div align="left">
				<input name="Save" type="submit" id="Save" value=" 修 改 " class="InputBox">
			</div>
			</td>
		</tr>
	</form>
</table>
<%
 	Call CloseRs
  Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" --></html>