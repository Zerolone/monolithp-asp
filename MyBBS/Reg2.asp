<!-- #include file ="../Inc/Monolith_Conn.asp" -->
<!-- #include file="Head.asp"-->
<!-- #include file="Skin.asp"--><%
	Dim Board_ID, Board_Name
  Board_ID 		= Request("Board_ID")
  Board_Name	= Request("Board_Name")
  
  Dim TheString
%>
<!-- #include file="LoginString.asp"-->
<!-- #include file="BoardString.asp"-->
<!-- #include file="WelcomeBack.asp"-->
<script language="javascript">
function Check()
{
	if (FrmReg.Users_Name.value == "")
	{
		alert("请输入您的用户名！");
		FrmReg.Users_Name.focus();
		return false;
	}

	if (FrmReg.Users_Password.value == "")
	{
		alert("请输入您的密码！");
		FrmReg.Users_Password.focus();
		return false;
	}

	if (FrmReg.Users_Password1.value == "")
	{
		alert("请重复输入您的密码！");
		FrmReg.Users_Password1.focus();
		return false;
	}

	if (FrmReg.Users_Password.value != FrmReg.Users_Password1.value)
	{
		alert("两次输入的密码不同， 请重新输入！");
		FrmReg.Users_Password.focus();
		return false;
	}

	if (FrmReg.Users_Email.value == "")
	{
		alert("请输入您的邮箱！");
		FrmReg.Users_Email.focus();
		return false;
	}

	if (FrmReg.Users_Password.value == "")
	{
		alert("请再次输入您的邮箱！");
		FrmReg.Users_Password.focus();
		return false;
	}

	if (FrmReg.Users_Password.value == "")
	{
		alert("请输入您的密码！");
		FrmReg.Users_Password.focus();
		return false;
	}


}
//-->
</script>
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><table width="100%" cellspacing="0" cellpadding="0" style="border: 1px solid #000000">
	<form method="POST" action="Reg3.asp" name="FrmReg" onsubmit="return Check()">
		<tr>
			<td background="images/tile_back.gif" height="32" colspan="2">&nbsp;<img height="8" src="images/nav_m.gif" width="8" border="0">
			<b><font color="#FFFFFF">开始注册本论坛注册</font></a></b></td>
		</tr>
		<tr>
			<td height="130" align="center" colspan="2">
			<table border="0" cellspacing="0" cellpadding="0" style="border: 1px solid #992A2A" width="99%">
				<tr>
					<td bgcolor="#E3C0C0" height="27"><b>&nbsp;<font color="#992A2A">Attention!
					</font></b></td>
				</tr>
				<tr>
					<td bgcolor="#F2DDDD" height="90">
					<p><font color="#992A2A">&nbsp;请确定你已经同意本论坛的条款</font></p>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="27%" class="pformleft">&nbsp;请输入你的名称</td>
			<td width="73%" class="pformright">
			<input class="InputBox" type="text" name="Users_Name" style="height:18; width=160"></td>
		</tr>
		<tr>
			<td width="27%" class="pformleft">&nbsp;请输入你的密码</td>
			<td width="73%" class="pformright">
			<input class="InputBox" type="password" name="Users_Password" style="height:18; width=160"></td>
		</tr>
		<tr>
			<td width="27%" class="pformleft">&nbsp;请输入重复密码</td>
			<td width="73%" class="pformright">
			<input class="InputBox" type="Password" name="Users_Password1" style="height:18; width=160"></td>
		</tr>
		<tr>
			<td width="27%" class="pformleft">&nbsp;请输入你的邮箱</td>
			<td width="73%" class="pformright">
			<input class="InputBox" type="text" name="Users_Email" style="height:18; width=160"></td>
		</tr>
		<tr>
			<td width="27%" class="pformleft">&nbsp;请输入重复邮箱</td>
			<td width="73%" class="pformright">
			<input class="InputBox" type="text" name="Users_Email1" style="height:18; width=160"></td>
		</tr>
		<tr>
			<td class="pformstrip" align="center" colspan="2">
			<input class="InputBox" type="submit" value=" 注 册 新 用 户 " name="submit"></td>
			</td>
		</tr>
	</form>
</table>
<!-- #include file="CopyRight.asp"-->