<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>�ޱ����ĵ�</title>
</head>

<script language="javascript" src="../Js/All.js"></script>

<form action=PayType.asp method=post name=paytype>
<table width="95%" border="1"  style="border-collapse: collapse;border:dotted 1px" cellspacing="2" cellpadding="2" align="center">
	<tr>
		<td width=100 align=right>֧����� </td>
		<td> <input type=text name=paytype class="InputBox"></td>
	</tr>
	<tr>
		<td align=right>��ʾ��Ϣ˵��</td>
		<td><textarea name="paymentmessage" rows="4" cols="55" class="InputBox"></textarea></td>
	</tr>
	<tr>
		<td align=right>���</td>
		<td>
			<input type="radio" name="paymark" value=1>������֧�� 
			<input type="radio" name="paymark" checked value=0>������֧��
		</td>
	</tr>
	<tr>
		<td align=right>����֧��URL</td>
		<td><input type=text name=payurl size=60 class="InputBox"></td>
	</tr>
	<tr>
		<td colspan=2 align=right>
			<input name=action type="submit" value=" �� �� "  class="InputBox">
		</td>
	</tr>
</table>
</form>