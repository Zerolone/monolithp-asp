<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
</head>
<script language="javascript" src="/Js/All.js"></script>
<script language="JavaScript">
<!--
function ConfirmDelBig()
{
   if(confirm("确定要删除此大类吗？删除此大类同时将删除所包含的小类和该类下的所有产品，并且不能恢复！"))
     return true;
   else
     return false;
	 
}

function ConfirmDelMid()
{
   if(confirm("确定要删除此中类吗？删除此中类同时将删除该类下的所有产品，并且不能恢复！"))
     return true;
   else
     return false;
}
-->
</script>
<%
	Dim ProductClass_BigName, ProductClass_BigNameTmp

	'----------------------0-----------------------1-----------------------2---------------------3
  sql="Select [ProductClass_BigOrder], [ProductClass_BigName], [ProductClass_MidOrder], [ProductClass_MidName]  from [Monolith_ProductClass] order by [ProductClass_BigName] Asc,  [ProductClass_BigOrder] asc , [ProductClass_MidOrder] asc"
  rs.open sql,conn,1,3
	if not (rs.bof or rs.eof) then
	  ProductClass_BigName=rs(1)
%>  
	<table width="100%" border="0" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="2">
		<b><font color="#FFFFFF">&nbsp;产品管理 &gt;&gt; 产品类别管理</font></b></td>
	</tr>
<%
  	Do While Not (rs.eof Or Rs.Bof)
    	if (ProductClass_BigName<>ProductClass_BigNameTmp) then
%>
  <tr height="22">
    <td width="83%" bgcolor="#FFFFFF">&nbsp;<%=rs(0)%><a href="?action=editclass&ProductClass_BigName=<%=rs(1)%>"><%=rs(1)%></a></td>
        <td width="16%" align="right" bgcolor="#FFFFFF"> ｜<a href="ClassAddMid.asp?ProductClass_BigName=<%=rs(1)%>" target="tabWin" onclick="fnClick();">增加中类</a>｜<a href="ClassModifyBig.asp?ProductClass_BigName=<%=rs(1)%>"  target="tabWin" onclick="fnClick();">修改</a>｜<a href="ClassDelBig.asp?ProductClass_BigName=<%=rs(1)%>"  onClick="return ConfirmDelBig();">删除</a>｜ 
        </td>
  </tr>
<%
      	ProductClass_BigName=rs(1)
			end if
%>
  <tr height="22">
    <td>&nbsp;&nbsp;<a href="admin_prod.asp?action=editclass&ProductClass_MidName=<%=rs(3)%>"><%=rs(2)%>-<%=rs(3)%></a></td>
    <td align="right">
	  ｜<a href="ClassModifyMid.asp?ProductClass_MidName=<%=rs(3)%>&ProductClass_BigName=<%=rs(1)%>"  target="tabWin" onclick="fnClick();">修改</a>｜<a href="ClassDelMid.asp?ProductClass_MidName=<%=rs(3)%>&ProductClass_BigName=<%=rs(1)%>" onClick="return ConfirmDelMid();">删除</a>｜
	</td>
 </tr>
<%
    	rs.movenext
    	if Not (Rs.Bof or Rs.Eof) then	ProductClass_BigNameTmp = rs(1)
  	Loop
	end if
%>
</table>
<% 
  Call CloseRs
  Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->