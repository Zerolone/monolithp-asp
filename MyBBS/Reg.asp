<!-- #include file ="../Inc/Monolith_Conn.asp" -->
<script language="JavaScript">
<!--
function Check()
{
	if (FrmReg.agree_cbox.checked == false)
	{
		alert("�������Ҫע�᱾��̳�� ��ͬ�Ȿ��̳��Э�飡");
		return (false);
	}
}
//-->
</script>
<!-- #include file="Head.asp"-->
<!-- #include file="Skin.asp"--><%
  Dim Board_ID, Board_Name
  Board_ID 		= Request("Board_ID")
  Board_Name	= Request("Board_Name")
  
  Dim TheString

  Dim BBS_ID 
  BBS_ID			= Request("BBS_ID")
%>
<!-- #include file="BoardString.asp"-->
<!--�ָ���--><img src="images/Blank.gif" border="0" width="1" height="5"><!--�ָ���--><table width="100%" cellspacing="0" cellpadding="0" style="border: 1px solid #000000">
<form method="GET" action="Reg2.asp" name="FrmReg" onsubmit="return Check()">
	<tr>
		<td background="images/tile_back.gif" height="32">
		<img height="8" src="images/nav_m.gif" width="8" border="0"> <b>
		<font color="#FFFFFF">ע������</font></b> </td>
	</tr>
	<tr>
		<td align="center" width="100%" background="images/tile_sub.gif" height="24" class="pformstrip">
		���������ͬ������Э��:</td>
	</tr>
	<tr>
		<td>
		<table cellspacing="1" cellpadding="4" width="100%" border="0">
			<tr>
				<td class="row4" width="100%" valign="top"><b>��̳�ŶӺ͹���</b> <br>
				<br>
				����1<br>
				<br>
				����2<br>
				<br>
				����3<br>
				<br>
				����4<br>
				<br>
				����5</td>
			</tr>			<tr>
				<td class="row4" width="100%" valign="top">
				<label for="agree_cbox">
				<input id="agree_cbox" type="checkbox" value="1" name="agree_to_terms">
				<b>���Ѿ��Ķ�, �˽Ⲣͬ�����ϵ���̳����</b></label></td>
			</tr>
			<tr>
				<td class="pformstrip" width="100%" align="center">
				<input class="InputBox" type="submit" value=" ע �� " name="submit"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#8394B2" height="5">
		<!--�ָ���--><img src="images/Blank.gif" border="0" width="1" height="5"><!--�ָ���--></td>
	</tr>
</form>
</table>
<!-- #include file="CopyRight.asp"-->

</body>

</html>