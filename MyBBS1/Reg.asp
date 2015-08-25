<!-- #include file="Inc/Conn.asp"-->
<html>
<head>
<title>新用户注册</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<script language="JavaScript">
<!--
function Showde(myform)
{

	if (myform.UserName.value == "")
	{
		alert("请输入您的登陆帐号！");
		myform.UserName.focus();
		return (false);
	}

	if (myform.Password.value == "")
	{
		alert("请输入您的登陆密码！");
		myform.Password.focus();
		return (false);
	}
	
	if (myform.Password.value != myform.Password1.value)
	{
		alert("请两次输入的密码不相同请重新输入！");
		myform.Password.value="";
		myform.Password1.value="";
		myform.Password.focus();
		return (false);
	}
	
if (document.all||document.getElementById){
for (i=0;i<myform.length;i++){
var tempobj=myform.elements[i]
if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
tempobj.disabled=true
}
}	
	
}
//-->
</script>
</head>
<body>
<form method="POST" action="RegAdd.asp" name="myform" onSubmit="return Showde(this)">
<table border="0" width="100%">
  <tr>
	<td width="131" align="right">登陆帐号：</td>
	<td><input type="text" name="UserName" size="32"></td>
  </tr>
  <tr>
	<td width="131" align="right">登陆密码：</td>
	<td><input type="password" name="Password" size="32"></td>
  </tr>
  <tr>
	<td width="131" align="right">重复密码：</td>
	<td><input type="password" name="Password1" size="32"></td>
  </tr>
</table>
<p align="center"><input type="submit" value=" 注 册 用 户" name="B1"></p>
</form>
<!-- #include file="BoardTime.asp"-->