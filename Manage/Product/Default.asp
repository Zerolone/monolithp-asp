<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Product_ID, action

  Product_ID = trim(Request("Product_ID"))
  action	= Request("Action")

  if action <> "" and Product_ID <> "" then
    if Action=" 删 除 " then
		  sql="delete from [Monolith_Product]	where [Product_ID] in (" & Product_ID & ")"
			conn.Execute sql
	  end if
	
		if Action = " 不 发 布 " then
		  Sql="update [Monolith_Product] set [Product_Online] = false  where [Product_ID] in (" & Product_ID & ")"
			conn.Execute sql
		end if

	  if Action = " 发 布 " then
		  Sql="update [Monolith_Product] set [Product_Online] = true  where [Product_ID] in (" & Product_ID & ")"
			conn.Execute sql
    end if
  end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
</head>

<script language="javascript" src="/Js/All.js"></script><script language="JavaScript">
<!--
function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
}
//-->
</script><%
	Dim ThisPage
  ThisPage					= Request.ServerVariables("PATH_INFO")
%>
<table width="100%" border="0" cellpadding="2" cellspacing="2">
	<form action="<%=ThisPage%>" name="search" method="get">
		<tr>
			<td>
			类别 <select name="Product_ClassMidName" class="InputBox">
			<option value selected>---所有类别---</option>
			<%
		  Sql = "Select Distinct [ProductClass_BigName],[ProductClass_BigOrder],[ProductClass_MidName],[ProductClass_MidOrder] From [Monolith_ProductClass] Order By [ProductClass_BigOrder],[ProductClass_MidOrder] " 
		  Rs.Open Sql,conn,1,3
		  if Rs.bof and Rs.eof then
		    response.write "<option Selected>---暂无分类---</option>"
		  else
			Do While Not (Rs.eof or Rs.Eof)
			  Response.write "<option value=""" & Rs("ProductClass_MidName")& """>" & Rs("ProductClass_BigName") & "-" & Rs("ProductClass_MidName") & "</option>"
			  Rs.movenext
			loop
		  end if
		  Rs.close
		%>
			</select> 关键字 <input type="text" name="keyword" size="10" maxlength="20" class="InputBox"> <input type="submit" value="搜索" class="InputBox"></td>
		</tr>
	</form>
</table>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="8">
		<b><font color="#FFFFFF">&nbsp;产品管理</font></b></td>
	</tr>
	<form name="Product_List" action="Default.asp" method="post">
		<tr height="20" bgcolor="#999999">
			<td>
			　</td>
			<td width="21%">
			<font color="#FFFFFF">&nbsp;编号(编辑)</font></td>
			<td width="19%">
			<font color="#FFFFFF">&nbsp;名称(预览)</font></td>
			<td width="13%">
			<font color="#FFFFFF">&nbsp;产品型号</font></td>
			<td width="12%">
			<font color="#FFFFFF">&nbsp;大类</font></td>
			<td width="10%">
			<font color="#FFFFFF">&nbsp;中类</font></td>
			<td width="13%">
			<font color="#FFFFFF">&nbsp;上线日期</font></td>
			<td width="7%">
			<font color="#FFFFFF">&nbsp;操作</font></td>
		</tr>
		<%
		dim ProductClass_MidName, KeyWord, actiontext
		Dim Page, Pages, j
	  ThisPage							= Request.ServerVariables("PATH_INFO")
		Page									= clng(request("Page"))
		Dim i, TrBgColor

    ProductClass_MidName	= Request("ProductClass_MidName")
    KeyWord					= Request("KeyWord")

		'-------------------0--------------1--------------2--------------3------------------4------------------5--------------------6----------------7
    Sql = "Select [Product_ID], [Product_Code], [Product_Name], [Product_Model], [Product_BigName], [Product_MidName], [Product_AddDate], [Product_Online] From [Monolith_Product] Where 1=1 "

	if ProductClass_MidName <> ""	then Sql = Sql & " And [ProductClass_MidName]='" & ProductClass_Name & "'"
	if KeyWord <> ""							then Sql = Sql & " And [Product_Name] like '%" & KeyWord & "%'  Or [Product_Model] like '%" & keyword & "%'"

	Rs.Open Sql, Conn ,1, 1
	if not (rs.bof or rs.eof) then
    rs.PageSize=12
		if page=0 then page=1
		pages=rs.pagecount
		if page > pages then page=pages
		rs.AbsolutePage=page  
		for j=1 to rs.PageSize 
		
				if j mod 2 = 0 then 
					TrBgColor	= ""
				else 
					TrBgColor	= "bgcolor=""#FFFFFF"""
				end if  %> <tr <%=trbgcolor%> height="22">
			<td>
			<input type="checkbox" value="<%=rs(0)%>" name="Product_ID"></td>
			<td width="21%">
			<a href="Modify.asp?Product_ID=<%=rs(0)%>&Page=<%=Page%>">&nbsp;<%=rs(1)%></a></td>
			<td width="19%">
			<a href="../../Product/ProductView.asp?Product_ID=<%=rs(0)%>" target="_blank">&nbsp;<%=rs(2)%></a></td>
			<td width="13%">
&nbsp;<%=rs(3)%></td>
			<td width="12%">
&nbsp;<%=rs(4)%></td>
			<td width="10%">
&nbsp;<%=rs(5)%></td>
			<td width="13%">
&nbsp;<%=rs(6)%></td>
			<%
	  	if rs(7)=true then
				Action		= " 不 发 布 "
				ActionText	= "不发布"
	  	else
				Action		= " 发 布 "
				ActionText	= "发布"
	  	end if
	%> <td width="7%">
&nbsp;<a href="Default.asp?action=<%=Action%>&Product_ID=<%=rs(0)%>"><%=ActionText%></a></td>
		</tr>
		<%
      	rs.movenext
	  		if rs.eof then exit for
				next

		Dim ColorMod
		
		if TrBgColor="" then
			ColorMod=0
		else
			ColorMod=1
		end if	

		if j < Rs.PageSize then
		  for i=1 to Rs.PageSize-j
				if i mod 2 = ColorMod then 
					TrBgColor	= ""
				else 
					TrBgColor	= "bgcolor=""#FFFFFF"""
				end if
	%> <tr <%=trbgcolor%> height="22">
			<td width="3%">
			　</td>
			<td width="21%">
			　</td>
			<td width="19%">
			　</td>
			<td width="13%">
			　</td>
			<td width="12%">
			　</td>
			<td width="10%">
			　</td>
			<td width="13%">
			　</td>
			<td width="7%">
			　</td>
		</tr>
		<%	
			next
		end if
		if TrBgColor="" then
			TrBgColor	= "bgcolor=""#FFFFFF"""
		else
			TrBgColor = ""
		end if
	%>
	 <tr <%=trbgcolor%>>
			<td colspan="8" align="center" height="28">
&nbsp;&nbsp;&nbsp; <input name="chkAll" type="checkbox" id="chkAll" onclick="CheckAll(this.form)" value="checkbox"> 全 选 <input type="submit" name="action" value=" 删 除 " onclick="{if(confirm('该操作不可恢复！\n\n确定删除选定的产品？')){this.document.Product_List.submit();return true;}return false;}" class="InputBox">&nbsp;&nbsp;
			<input type="submit" name="action" value=" 不 发 布 " onclick="{if(confirm('确定不发布选定的产品在线？')){this.document.Product_List.submit();return true;}return false;}" class="InputBox">&nbsp;&nbsp;
			<input type="submit" name="action" value=" 发 布 " onclick="{if(confirm('确定发布选定的产品在线？')){this.document.Product_List.submit();return true;}return false;}" class="InputBox"> </td>
		</tr>
	</form>
	<%
		if TrBgColor="" then
			TrBgColor	= "bgcolor=""#FFFFFF"""
		else
			TrBgColor = ""
		end if
%> <tr <%=trbgcolor%> height="30" valign="bottom">
		<form method="Post" action="<%=ThisPage%>" style="MARGIN-BOTTOM:0px">
			<td colspan="8" align="center">
			<%
			  ThisPage = ThisPage & "?1=1"
			  if Page<2 then
			%> ｜ <strike>首页</strike> ｜ <strike>上一页</strike> ｜ <%else%> ｜ <a href="<%=ThisPage%>&page=1">首页</a> ｜ <a href="<%=ThisPage%>&page=<%=Page-1%>">上一页</a> ｜ <%
				end if
				if rs.pagecount-page<1 then
			%> <strike>下一页</strike> ｜ <strike>尾页</strike> ｜ <%else%> <a href="<%=ThisPage%>&page=<%=page+1%>">下一页</a> ｜ <a href="<%=ThisPage%>&page=<%=rs.pagecount%>">尾页</a> ｜ <%end if%> 页次：<strong><font color="red"><%=Page%></font>/<%=rs.pagecount%></strong> 页 ｜ 共<b><font color="#FF0000"><%=rs.recordcount%></font></b>条记录 
			｜ <b><%=rs.pagesize%></b>条记录/页 ｜ 转到：<input type="text" name="page" size="4" maxlength="10" class="InputBox" value="<%=page%>"> <input class="InputBox" type="submit" value=" Goto " name="cndok"> <%
	Call CloseRs
	Call CloseConn
				end if

%> </td>
		</form>
	</tr>
</table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->

</html>