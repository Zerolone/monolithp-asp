<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>�ޱ����ĵ�</title>
</head>

<script language="javascript" src="../Js/All.js"></script>
<%
	Dim Action, PayDefault_ID, m

	Action = Request("Action")
	PayDefault_ID =  Request("Paydefault_id")

	if action=" �� �� " then
		sql="Select * from [Monolith_Paydefault]"
		rs.open sql,conn,1,3
		Rs.AddNew
		rs("Paydefault_paytype")=Request.Form("paytype")
		rs("Paydefault_paymentmessage")=Request.Form("paymentmessage")
		rs("Paydefault_paymark")=Request.Form("paymark")
		rs("Paydefault_payurl")=Request.Form("payurl")
    rs.update
		rs.close
	end if

	if action=" �� �� " then
		sql="Select * from [Monolith_Paydefault] where Paydefault_ID="& PayDefault_ID
		rs.open sql,conn,1,3
		rs("Paydefault_paytype")=Request.Form("paytype")
		rs("Paydefault_paymentmessage")=Request.Form("paymentmessage")
		rs("Paydefault_paymark")=Request.Form("paymark")
		rs("Paydefault_payurl")=Request.Form("payurl")
    rs.update
		rs.close
	end if

	if action=" ɾ �� " then
    Dim StrSQL
    StrSQL="delete from [Monolith_PayDefault] where [PayDefault_ID]=" & PayDefault_ID
    conn.Execute StrSQL    
    conn.close
		Response.Redirect "Admin_Order_PayType.asp"
	end if

	Sql = "Select * From [Monolith_PayDefault]"
	Rs.open Sql, Conn ,1, 3
	m=1
	do while not rs.eof
%>
<form action=PayType.asp method=post name=paytype<%=m%>>
<table width="95%" border="1"  style="border-collapse: collapse;border:dotted 1px" cellspacing="2" cellpadding="2" align="center">
	<tr>
		<td width=100 align=right>֧����� <font color=red><b><%=m%></b></font></td>
		<td> <input type=text value="<%=rs("PayDefault_paytype")%>" name=paytype class="InputBox"></td>
	</tr>
	<tr>
		<td align=right>��ʾ��Ϣ˵��</td>
		<td><textarea name="paymentmessage" rows="4" cols="55" class="InputBox"><%=rs("PayDefault_paymentmessage")%></textarea></td>
	</tr>
	<tr>
		<td align=right>���</td>
		<td>
		<%if rs("PayDefault_paymark")="1" then %>
			<input type="radio" name="paymark" checked value=1>������֧�� 
			<input type="radio" name="paymark" value=0>������֧��
		<%else%>
			<input type="radio" name="paymark" value=1>������֧�� 
			<input type="radio" name="paymark" checked value=0>������֧��
		<%end if%>
		</td>
	</tr>
	<tr>
		<td align=right>����֧��URL</td>
		<td><input type=text value="<%=rs("PayDefault_payurl")%>" name=payurl size=60 class="InputBox"></td>
	</tr>
	<tr>
		<td colspan=2 align=right>
			<input name=action type="submit" value=" �� �� " class="InputBox">
			<input name=action type="submit" value=" ɾ �� " class="InputBox">
		</td>
	</tr>
		<input  type="hidden" name=PayDefault_id value=<%=rs("PayDefault_ID")%>>
</table>
</form>
<%
		rs.movenext 
		m=m+1
	loop 
	Call CloseRs
	Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->