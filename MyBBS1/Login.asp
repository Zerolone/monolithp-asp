<!-- #include file="Inc/Conn.asp"-->
<script language="javascript">
function Showde(myform)
{
	if (myform.UserName.value == "")
	{
		alert("�����������û�����");
		myform.UserName.focus();
		return (false);
	}

	if (myform.Password.value == "")
	{
		alert("�������������룡");
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
    <td width="131" align="right">&nbsp;�û�����</td>
	<td><input type="text" name="UserName" size="32"></td>
  </tr>
  <tr>
	<td width="131" align="right">&nbsp;��&nbsp; �룺</td>
	<td><input type="password" name="Password" size="32"></td>
  </tr>
  <tr>
    <td width="131" align="right">�������ڣ�</td>
	<td><input type="radio" value="0" name="CookiesTime">������<input type="radio" value="1" name="CookiesTime">һ��<input type="radio" value="7" name="CookiesTime">һ��<input type="radio" value="30" name="CookiesTime" checked>һ����<input type="radio" value="365" name="CookiesTime">һ��</td>
  </tr>
</table>
<p align="center">	&nbsp;<input type="submit" value=" �� ½ " name="B1"></p>
</form>
<!-- #include file="BoardTime.asp"-->