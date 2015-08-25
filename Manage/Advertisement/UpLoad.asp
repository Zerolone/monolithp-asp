<%
  Option Explicit

  Response.Buffer = true
  Response.ExpiresAbsolute=now()-1
  Response.Expires=0
  Response.CacheControl="no-cache"
%>
<!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>图片上传</title>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>
<script language="javascript">
function AddTo()
{
  window.returnValue = document.form1.Advertisement_FileName.value;
  window.close();
  
//  document.form1.Submit_AddFileName.click();
}
</script>
<body>
<form name="form1" method="post" action="UpFile.asp" enctype="multipart/form-data" target="AAA">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2">
			<b><font color="#FFFFFF">&nbsp;图片上传</font></b></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" align="right" height="20">
			<font color="#FFFFFF">文&nbsp;&nbsp;&nbsp; 件：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="file" name="file2" size="38" class="InputBox"> </td>
		</tr>
		<tr>
			<td bgcolor="#999999" align="right" height="20" width="100">
			<font color="#FFFFFF">类&nbsp;&nbsp;&nbsp; 型：</font></td>
			<td bgcolor="#FFFFFF">
			&nbsp;Jpg, Gif, Swf 限制：200K 			</td>
		</tr>
		<tr>
			<td bgcolor="#999999" height="20" width="100">
			　</td>
			<td bgcolor="#FFFFFF">
			<input type="submit" name="Submit" value=" 上 传 文 件 " class="InputBox">
			<input type="hidden" name="Submit_AddFileName" onclick="AddTo();" style="display=none;">
			<input type="hidden" name="Advertisement_FileName">
			</td>
		</tr>
	</table>
</form>
<iframe name="AAA" width="0" height="0" style="display=none;"></iframe>

</body>

</html>