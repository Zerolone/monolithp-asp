<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript">
function check()
{
  var Check_Bulletin_Title	= document.FrmAddMenu.Bulletin_Title;
  if (Check_Bulletin_Title.value == "")
  {
 	alert("�����빫����⣡");
	Check_Bulletin_Title.focus();
	return false;
  }
  
  var Check_Bulletin_Content	= document.FrmAddMenu.Bulletin_Content;
  if (Check_Bulletin_Content.value == "")
  {
 	alert("�����빫�����ݣ�");
	Check_Bulletin_Content.focus();
	return false;
  }
}
</script>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<form method="POST" action="Add_Submit.asp" onsubmit="return check()" name="FrmAddMenu">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;������� &gt;&gt; �������</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� ��&nbsp;&nbsp;&nbsp;&nbsp; ID��</font></td>
			<td bgcolor="#FFFFFF">ϵͳ�Զ�����</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� ��&nbsp; �� �⣺</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Bulletin_Title" size="40" style="width=250px;" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� �� �����⣺</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Bulletin_Title2" size="40" style="width=250px;" class="InputBox">
			</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�����������ݣ�</font></td>
			<td bgcolor="#FFFFFF"><font color="#FF0000">
			<textarea rows="2" name="Bulletin_Content2" cols="20" style="width=500;height:100" class="InputBox"></textarea></font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">��һ��Ϊ�������ݣ�</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� ��&nbsp; �� �ݣ�</font></td>
			<td bgcolor="#FFFFFF"><font color="#FF0000">
			<textarea rows="2" name="Bulletin_Content" cols="20" style="width=500;height:100" class="InputBox"></textarea> 
			*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">��������ʾ���ݣ�</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">
			<input type="submit" value=" �� �� (Alt + S) " name="B1" class="InputBox" accesskey="S">
			<input type="reset" value=" �� �� (Alt + R) " name="B2" class="InputBox" accesskey="R"></td>
		</tr>
	</form>
</table>
<table><tr><td height="144"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->