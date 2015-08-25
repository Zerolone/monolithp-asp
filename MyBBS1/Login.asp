<!-- #include file="Inc/Conn.asp"-->
<script language="javascript">
function Showde(myform)
{
	if (myform.UserName.value == "")
	{
		alert("请输入您的用户名！");
		myform.UserName.focus();
		return (false);
	}

	if (myform.Password.value == "")
	{
		alert("请输入您的密码！");
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

<form method="POST" action="LoginCheck.asp" name="myform" onSubmit="return Showde(this)">
<table border="0" width="100%">
  <tr>
    <td width="131" align="right">&nbsp;用户名：</td>
	<td><input type="text" name="UserName" size="32"></td>
  </tr>
  <tr>
	<td width="131" align="right">&nbsp;密&nbsp; 码：</td>
	<td><input type="password" name="Password" size="32"></td>
  </tr>
  <tr>
    <td width="131" align="right">保存日期：</td>
	<td><input type="radio" value="0" name="CookiesTime">不保存<input type="radio" value="1" name="CookiesTime">一天<input type="radio" value="7" name="CookiesTime">一周<input type="radio" value="30" name="CookiesTime" checked>一个月<input type="radio" value="365" name="CookiesTime">一年</td>
  </tr>
</table>
<p align="center">	&nbsp;<input type="submit" value=" 登 陆 " name="B1"></p>
</form>
<!-- #include file="BoardTime.asp"-->