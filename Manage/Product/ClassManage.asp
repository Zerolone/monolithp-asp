<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>�ޱ����ĵ�</title>
</head>
<script language="javascript" src="/Js/All.js"></script>
<script language="JavaScript">
<!--
function ConfirmDelBig()
{
   if(confirm("ȷ��Ҫɾ���˴�����ɾ���˴���ͬʱ��ɾ����������С��͸����µ����в�Ʒ�����Ҳ��ָܻ���"))
     return true;
   else
     return false;
	 
}

function ConfirmDelMid()
{
   if(confirm("ȷ��Ҫɾ����������ɾ��������ͬʱ��ɾ�������µ����в�Ʒ�����Ҳ��ָܻ���"))
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
		<b><font color="#FFFFFF">&nbsp;��Ʒ���� &gt;&gt; ��Ʒ������</font></b></td>
	</tr>
<%
  	Do While Not (rs.eof Or Rs.Bof)
    	if (ProductClass_BigName<>ProductClass_BigNameTmp) then
%>
  <tr height="22">
    <td width="83%" bgcolor="#FFFFFF">&nbsp;<%=rs(0)%><a href="?action=editclass&ProductClass_BigName=<%=rs(1)%>"><%=rs(1)%></a></td>
        <td width="16%" align="right" bgcolor="#FFFFFF"> ��<a href="ClassAddMid.asp?ProductClass_BigName=<%=rs(1)%>" target="tabWin" onclick="fnClick();">��������</a>��<a href="ClassModifyBig.asp?ProductClass_BigName=<%=rs(1)%>"  target="tabWin" onclick="fnClick();">�޸�</a>��<a href="ClassDelBig.asp?ProductClass_BigName=<%=rs(1)%>"  onClick="return ConfirmDelBig();">ɾ��</a>�� 
        </td>
  </tr>
<%
      	ProductClass_BigName=rs(1)
			end if
%>
  <tr height="22">
    <td>&nbsp;&nbsp;<a href="admin_prod.asp?action=editclass&ProductClass_MidName=<%=rs(3)%>"><%=rs(2)%>-<%=rs(3)%></a></td>
    <td align="right">
	  ��<a href="ClassModifyMid.asp?ProductClass_MidName=<%=rs(3)%>&ProductClass_BigName=<%=rs(1)%>"  target="tabWin" onclick="fnClick();">�޸�</a>��<a href="ClassDelMid.asp?ProductClass_MidName=<%=rs(3)%>&ProductClass_BigName=<%=rs(1)%>" onClick="return ConfirmDelMid();">ɾ��</a>��
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