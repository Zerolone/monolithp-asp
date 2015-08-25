<!-- #include virtual ="/Inc/Monolith_Conn.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script event="onclick" for="Ok3" language="JavaScript">
  window.returnValue = UpPicSelectList.value;
  window.close();
</script>
</head>

<body>
<table id="tb" cellspacing="0" cellpadding="0" width="490" border="0" name="tb" style="border: 1px solid #C0C0C0">
	<tr>
		<td height="270">
		<iframe name="I1" src="UpLoadedPic_News_Pic.asp" marginwidth="1" marginheight="1" height="100%" width="100%" border="0" frameborder="0">
		浏览器不支持嵌入式框架，或被配置为不显示嵌入式框架。</iframe>
		</td>
	</tr>
	<tr>
	<td>
			<form name="form2" style="MARGIN-BOTTOM:0px">
			<input type="hidden" value="" name="UpPicSelectList">
			<input type="hidden" name="InsertUpLoadPic0" id="Ok3">
			</form>
			</td>
	</tr>
</table></body>

</html>
