<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
</head>

<script language="javascript" src="../Js/All.js"></script>

<script language="JavaScript">
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
</script>
<%
	Dim ThisPage, Order_List_OrderNum

  ThisPage							= Request.ServerVariables("PATH_INFO")
	Order_List_OrderNum = Request("Order_List_OrderNum")
	
	if Request("Action") = "Update" then
		Sql="Select * From [Monolith_Order_List] where [Order_List_OrderNum]='" & Order_List_OrderNum &"'"
		rs.Open sql,conn,1,3
		rs(8)			= Request("Order_List_Status")
		rs(7)				= Request("Order_List_Memo")
		rs.update
		rs.close
		set rs=nothing
		Response.Redirect ThisPage & "?Order_List_OrderNum=" & Order_List_OrderNum
	end if
		
	'------------------------0-------------------1-----------------------2----------------------3----------------------4--------------------5--------------------6-------------------7-------------------8
	Sql = "Select [Order_List_Recname], [Order_List_RecPhone], [Order_List_Recmail], [Order_List_RecAddress], [Order_List_Zipcode], [Order_List_Paytype], [Order_List_Notes], [Order_List_Memo], [Order_List_Status] From [Monolith_Order_List] Where [Order_List_OrderNum]='" & Order_List_OrderNum &"'"
	Rs.Open Sql, Conn, 1, 3
  if Rs.eof and Rs.bof  then
		response.write "<Table><Tr><Td>订单错误!</Td></Tr></Table>"
		response.end
	end if
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
<form action="<%=ThisPage%>?Action=Update&Order_List_OrderNum=<%=Order_List_OrderNum%>" method=post>
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="8">
		<b><font color="#FFFFFF">&nbsp;产品订单管理 &gt;&gt; 订单修改/察看</font></b></td>
	</tr>
	<tr>
		<td valign='top' align='right' bgcolor="#999999" height="20"><font color="#FFFFFF">订单号：</font></td>
		<td bgcolor="#FFFFFF"><b><font color=red> <%=Order_List_OrderNum%> </font></b></td>
	</tr>
	<tr>
		<td valign='top' align='right' bgcolor="#999999" height="20"><font color="#FFFFFF">收货人姓：</font></td>
		<td bgcolor="#FFFFFF">
		<input type="text" name="Recname" value="<%=rs(0)%>" disabled size="39" maxlength="20" class="InputBox"></td>
	</tr>
	<tr>
		<td width="11%"  valign='top' align='right' bgcolor="#999999" height="20">
		<font color="#FFFFFF">收货人电话：</font></td>
		<td bgcolor="#FFFFFF">
		<input type="text" name="RecPhone" value="<%=rs(1)%>" disabled size="39" maxlength="20" class="InputBox"></td>
	</tr>
	<tr>
		<td valign='top' align='right' bgcolor="#999999" height="20"><font color="#FFFFFF">收货人邮箱：</font></td>
		<td bgcolor="#FFFFFF">
		<input type="text" name="Recmail" value="<%=rs(2)%>" disabled size="39" maxlength="30" class="InputBox"></td>
	</tr>
	<tr>
		<td valign='top' align='right' bgcolor="#999999" height="20"><font color="#FFFFFF">收货地址：</font></td>
		<td bgcolor="#FFFFFF">
		<input type="text" name="RecAddress" value="<%=rs(3)%>" disabled size="39" maxlength="30" class="InputBox"></td>
	</tr>
	<tr>
		<td valign='top' align='right' bgcolor="#999999" height="20"><font color="#FFFFFF">邮政编码：</font></td>
		<td bgcolor="#FFFFFF"><input type="text" name="Zipcode" value="<%=rs(4)%>" disabled size="10" maxlength="10" class="InputBox"></td>
	</tr>
	<tr>
		<td valign='top' align='right' bgcolor="#999999" height="20"><font color="#FFFFFF">支付类型：</font></td>
		<td bgcolor="#FFFFFF"><%=rs(5)%></td>
	</tr>
	<tr>
		<td valign='top' align='right' bgcolor="#999999" height="20"><font color="#FFFFFF">顾客说明：</font></td>
		<td bgcolor="#FFFFFF"><textarea cols='50' rows='4' name='Notes' disabled class="InputBox"><%=rs(6)%></textarea></td>
	</tr>
	<tr>
		<td valign='top' align='right' width="11%" bgcolor="#999999" height="20">
		<font color="#FFFFFF">处理备忘：</font></td>
		<td bgcolor="#FFFFFF"><textarea name='Order_List_Memo' cols='30' rows='4' id="Order_List_Memo" class="InputBox"><%=rs(7)%></textarea> 订单处理附加说明</td>
	</tr>
	<tr>
		<td align=right valign=top bgcolor="#999999" height="20"><font color="#FFFFFF">状态修改：</font></td>
		<td bgcolor="#FFFFFF">
			<select name="Order_List_Status">
			  <option value="<%=rs(8)%>"><%=rs(8)%></option>
			  <option value="未处理">未处理</option>
			  <option value="已处理">已处理</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="#999999" height="20"></td>
		<td bgcolor="#FFFFFF">
			<input type='submit' name='action' value='修改定单' class="InputBox">
			<input type='reset' name='reset' value='重设' class="InputBox">
		</td>
	</tr>
</form>
</table>
<%Rs.close%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="8">
		<b><font color="#FFFFFF">&nbsp;产品订单管理 &gt;&gt; 订购产品察看</font></b></td>
	</tr>
		<tr heigt="22">
		<td width="25%" height="22" bgcolor="#999999"><font color="#FFFFFF">&nbsp;产品名称</font></td>
		<td width="25%" height="22" bgcolor="#999999"><font color="#FFFFFF">&nbsp;购买数量</font></td>
		<td width="25%" height="22" bgcolor="#999999"><font color="#FFFFFF">&nbsp;购买单价</font></td>
		<td width="25%" height="22" bgcolor="#999999"><font color="#FFFFFF">&nbsp;合 计</font></td>
	</tr>
	<%
		Dim Total, Sum, j, TrBgColor
		'-----------------------0------------------------1-------------------------2--------------------------3
		Sql = "Select [Order_Detail_ProdUnit], [Order_Detail_BuyPrice], [Order_Detail_Product_ID], [Order_Detail_Product_Name] From [Monolith_Order_Detail] Where [Order_Detail_OrderNum] = '" & Order_List_OrderNum & "'"
		Rs.Open Sql, Conn , 1, 3
		Total = 0
	 	do while not Rs.eof 
	 		Sum = rs(0)* clng(rs(1))
	 		Sum = FormatNumber(Sum,2) 
	 		Total = Sum + Total '计算总金额
'	 		Discount = Rs("Discount") '因为只有订单列表库中有折扣，而且订单只有一条记录，所以必须先取出来暂存
			if j mod 2 = 0 then 
				TrBgColor	= ""
			else 
				TrBgColor	= "bgcolor=""#FFFFFF"""
			end if	
	%>
		<tr <%=trbgcolor%> heigt="22">
				<td width="25%"><a href="../Product/ProductView.asp?Product_ID=<%=rs(2)%>" target="_blank">&nbsp;<%=rs(3)%></a></td>
		<td width="25%">&nbsp;<%=rs(0)%></td>
		<td width="25%">&nbsp;<%=FormatNumber(rs(1),2)%></td>
		<td width="25%">&nbsp;￥<%=Sum%></td>
	</tr>
	<%
			Rs.movenext
			j = j + 1
		loop
	Call CloseRs
	Call CloseConn
	%>
</table>	
	
	<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->