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
	Dim Order_List_Status, Order_List_OrderTime, Order_List_Men, ThisPage, Page, Pages, i, j, TrBgColor

	Dim Order_List_OrderNum, action
  Order_List_OrderNum= trim(Request("Order_List_OrderNum"))
  action	= Request("Action")

  if Order_List_OrderNum <> "" then
    if Action=" ɾ �� " then
	  sql="delete from [Monolith_Order_List]		where [Order_List_OrderNum] in (" & Order_List_OrderNum & ")"
	  conn.Execute sql
	  sql="delete from [Monolith_Order_Detail]	where [Order_Detail_OrderNum] in (" & Order_List_OrderNum & ")"
	  conn.Execute sql
    end if
  end if

	Order_List_Status			= Request("Order_List_Status")
	Order_List_OrderTime 	= Request("Order_List_OrderTime")
	Order_List_Men				= Request("Order_List_Men")

  ThisPage							= Request.ServerVariables("PATH_INFO")
	page									= clng(request("page"))
%>
<table width=100% border=0 cellpadding=2 cellspacing=2>
<form action="<%=ThisPage%>" name="search" method=get>
  <tr>
	  <td>״̬
		  <select name="Order_List_Status" id="Order_List_Status" class="InputBox">
			  <option value="" selected>---���ж���---</option>
				<%
				  sql="Select Distinct [Order_List_Status] from [Monolith_Order_List]" 
					rs.Open sql,conn,1,3
					Do While Not rs.eof
					  response.write "<option value='" & rs(0) & "'>" & rs(0) & "</option>"
					  rs.movenext
					loop
					rs.close
				%>
			</select>
		</td>
		<td>����/�ջ��� 
		  <input type="text" name="Order_List_Men" size="17" maxlength="20" class="InputBox"></td>
		<td>
			<input type="radio" name="Order_List_OrderTime" value="yesterday">����
			<input type="radio" name="Order_List_OrderTime" value="today">����
			<input type="radio" name="Order_List_OrderTime" value="" checked>����
		</td>
		<td><input type="submit" value=" �� �� " class="InputBox"></td>
	</tr>
	</form>
</table>
	<script language="JavaScript">
	<!--
	function checkdel(delid){	
	if(confirm('ɾ��ѡ���Ķ���'))
	{location.href="admin_prod.asp?action=del&id="+delid;}
	}
	//-->
	</script>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
<form name=Order_List action=Default.asp method=post>
	<tr bgcolor="#6A6A6A" height="26">
		<td colspan="8">
		<b><font color="#FFFFFF">&nbsp;��Ʒ��������</font></b></td>
	</tr>
  <tr heiht="20">	
		<td bgcolor="#999999" height="20"><font color="#FFFFFF">&nbsp;�������</font></td> 
		<td  bgcolor="#999999" height="20"><font color="#FFFFFF">&nbsp;�ܽ��</font></td>
		<td bgcolor="#999999" height="20"><font color="#FFFFFF">&nbsp;֧������</font></td>
		<td bgcolor="#999999" height="20"><font color="#FFFFFF">&nbsp;�ջ���</font></td>	
		<td widdth=60 bgcolor="#999999" height="20"><font color="#FFFFFF">&nbsp;�µ�ʱ��</font></td>
		<td bgcolor="#999999" height="20"><font color="#FFFFFF">&nbsp;״̬</font></td>
		<td bgcolor="#999999" height="20">��</td>
	</tr>
<%
	Sql = "Select [Order_List_OrderNum], [Order_List_OrderSum], [Order_List_PayType], [Order_List_RecName], [Order_List_OrderTime], [Order_List_Status] From [Monolith_Order_List] where [Order_List_Del]=false"
	if Order_List_Status <> "" 					then Sql = Sql + " And [Order_List_Status]='" & Order_List_Status & "'"
	if Order_List_OrderTime="today" 		then Sql = Sql + " And [Order_List_OrderTime] = Date()"
	if Order_List_OrderTime="yesterday" then Sql = Sql + " And [Order_List_OrderTime] = Date()-1"
	if Order_List_Men<>"" 							then Sql = Sql + " And [Order_List_RecName] like '%" & Order_List_Men & "%' Or [Order_List_UserName] like '%" & Order_List_Men & "%'"
	Sql = Sql+" Order By [Order_List_OrderTime] Desc"
'	Response.write sql
	Rs.Open Sql, Conn ,1, 3
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
			end if		
	%>
  <tr <%=trbgcolor%> height="22">
	  <td>&nbsp;<a href="View.asp?Order_List_OrderNum=<%=rs(0)%>"  target="tabWin" onclick="fnClick();"><%=rs(0)%></a></td>
		<td>&nbsp;<%=rs(1)%></td>
		<td>&nbsp;<%=rs(2)%></td>
		<td>&nbsp;<%=rs(3)%></td>
		<td>&nbsp;<%=rs(4)%></td>
		<td>&nbsp;<%=rs(5)%></td>
		<td><input type="checkbox" value="'<%=rs(0)%>'" name=Order_List_OrderNum></td>
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
			<td width="22%">
			��</td>
			<td width="14%">
			��</td>
			<td width="17%">
			��</td>
			<td width="15%">
			��</td>
			<td width="13%">
			��</td>
			<td width="10%">
			��</td>
			<td width="6%">
			��</td>
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
	  <td colspan=6 class=b align=right>
		  <input type="submit" name="action" value=" ɾ �� " onClick="{if(confirm('�ò������ɻָ���\n\nȷʵɾ��ѡ���Ķ�����')){this.document.Order_List.submit();return true;}return false;}" class="InputBox"> 	
		</td>
		<td><input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox"></td>
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
	  ThisPage = ThisPage & "?Order_List_Status=" & Order_List_Status & "&Order_List_OrderTime=" & Order_List_OrderTime & "&Order_List_Men=" & Order_List_Men
			  if Page<2 then
			%> �� <strike>��ҳ</strike> �� <strike>��һҳ</strike> �� <%else%> �� <a href="<%=ThisPage%>&page=1">��ҳ</a> �� <a href="<%=ThisPage%>&page=<%=Page-1%>">��һҳ</a> �� <%
				end if
				if rs.pagecount-page<1 then
			%> <strike>��һҳ</strike> �� <strike>βҳ</strike> �� <%else%> <a href="<%=ThisPage%>&page=<%=page+1%>">��һҳ</a> �� <a href="<%=ThisPage%>&page=<%=rs.pagecount%>">βҳ</a> �� <%end if%> ҳ�Σ�<strong><font color="red"><%=Page%></font>/<%=rs.pagecount%></strong> ҳ �� ��<b><font color="#FF0000"><%=rs.recordcount%></font></b>����¼ 
			�� <b><%=rs.pagesize%></b>����¼/ҳ �� ת����<input type="text" name="page" size="4" maxlength="10" class="InputBox" value="<%=page%>"> <input class="InputBox" type="submit" value=" Goto " name="cndok"> <%
	Call CloseRs
	Call CloseConn
				end if

%> </td>
		</form>
	</tr>
</table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->