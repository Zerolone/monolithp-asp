<!-- #include file ="../Inc/Monolith_Conn.asp" -->
<%
  Response.cookies(Monolith_CookiesName)("Users_ID")				= ""
  Response.cookies(Monolith_CookiesName)("Users_Name")			= ""
  Response.cookies(Monolith_CookiesName)("Users_Password")	= ""
  Response.cookies(Monolith_CookiesName).expires						= date
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="refresh" content="3; url=Default.asp" />
<title>MyBBS (Power By Monolith)</title>
<link href="../Css/BBS.css" rel="stylesheet" type="text/css">
</head>

<body>

<table border="0" width="100%" height="100%">
	<tr>
		<td align="center">��<table width="50%" cellspacing="0" cellpadding="0" style="border: 1px solid #000000">
			<tr>
				<td class="pformstrip"><b>&nbsp;��л!</b></td>
			</tr>
			<tr>
				<td height="62" class="row4">
				<p>&nbsp;���Ѿ��˳�<br>
&nbsp;���Ժ�,����Ϊ��ת��ҳ��...<b> </b></p>
				</td>
			</tr>
			<tr>
				<td class="pformstrip" height="28" align="center">(<a href="Default.asp">����㲻��ȴ��밴����</a>) 
				��</td>
			</tr>
		</table>
		</td>
	</tr>
</table>

</body>

</html>
