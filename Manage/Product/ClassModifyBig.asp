<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
</head>

<%
	Dim ThisPage, Action , ProductClass_BigName, New_ProductClass_BigName, Old_ProductClass_BigName, FoundErr

  ThisPage					= Request.ServerVariables("PATH_INFO")
  Action					= trim(Request("Action"))
  ProductClass_BigName		= trim(Request("ProductClass_BigName"))
  New_ProductClass_BigName	= trim(Request("New_ProductClass_BigName"))
  Old_ProductClass_BigName	= trim(Request("Old_ProductClass_BigName"))

  if ProductClass_BigName="" then response.Redirect("Admin_Product_ClassManage.asp")
  Sql = "Select * from [Monolith_ProductClass] where [ProductClass_BigName]='" & ProductClass_BigName & "'"
  rs.Open sql,conn,1,3
  if rs.Bof and rs.EOF then
	Response.write "<br><li>此产品大类不存在！</li>"
  else
	if Action="Modify" then
	  if ProductClass_BigName="" then
		FoundErr=True
		Resonse.write "<br><li>产品大类名不能为空！</li>"
  	  end if
	  if FoundErr<>True then
		rs("ProductClass_BigName") = New_ProductClass_BigName
   	 	rs.update
		rs.Close
     	set rs=Nothing
		if New_ProductClass_BigName <> Old_ProductClass_BigName then
		  Sql = "Update [Monolith_ProductClass] 	set [ProductClass_BigName]='" & New_ProductClass_BigName & "' where [ProductClass_BigName]='" & Old_ProductClass_BigName & "'"
		  conn.execute(Sql)
		  Sql = "Update [Monolith_Product] 			set [Product_BigName]='" & New_ProductClass_BigName & "' where [Product_BigName]='" & Old_ProductClass_BigName & "'"
		  conn.execute(Sql) 
     	end if		

	Response.write "<script language=""javascript"">"
	Response.write "  alert(""大类修改成功，请刷新页面。"");"
	Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
	Response.write "</script>"


			end if
	end if
	if FoundErr=True then
		call WriteErrMsg()
	else
%> <script language="JavaScript" type="text/JavaScript">
function checkBig()
{
  if (document.form1.ProductClass_BigName.value=="")
  {
    alert("大类名称不能为空！");
    document.form1.ProductClass_BigName.focus();
    return false;
  }
}
</script><table width="100%" border="0" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="2">
		<b><font color="#FFFFFF">&nbsp;产品管理 &gt;&gt; 产品类别管理 &gt;&gt; 产品大类修改</font></b></td>
	</tr>
	<tr>
		<td align="center" valign="top">
		<form name="form1" method="post" action="<%=ThisPage%>?Action=Modify">
		</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right" height="20">
			<font color="#FFFFFF">大类名称：</font></td>
			<td bgcolor="#E3E3E3">
			<input name="New_ProductClass_BigName" type="text" id="New_ProductClass_BigName" value="<%=rs("ProductClass_BigName")%>" size="20" maxlength="30" class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" bgcolor="#999999" width="100" height="20">
			　</td>
			<td align="center" bgcolor="#E3E3E3">
			<div align="left">
				<input name="ProductClass_BigName" type="hidden" id="ProductClass_BigName" value="<%=rs("ProductClass_BigName")%>"><input name="Old_ProductClass_BigName" type="hidden" id="Old_ProductClass_BigName" value="<%=rs("ProductClass_BigName")%>">
				<input name="Save" type="submit" id="Save" value=" 修 改 " class="InputBox"> </div>
			</td>
		</tr>
	</form>
</table>
</form>
</td>
</tr>
</table>
<%
	end if
end if
Call CloseRs
Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->