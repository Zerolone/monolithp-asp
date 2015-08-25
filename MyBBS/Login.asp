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
function Showde(myform)
{
	if (myform.Users_Name.value == "")
	{
		alert("请输入您的用户名！");
		myform.UserName.focus();
		return false;
	}

	if (myform.Users_Password.value == "")
	{
		alert("请输入您的密码！");
		myform.Password.focus();
		return false;
	}
}
//-->
</script>
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><table width="100%" cellspacing="0" cellpadding="0" style="border: 1px solid #000000">
	<form method="POST" action="LoginCheck.asp" name="myform" onsubmit="return Showde(this)">
		<tr>
			<td background="images/tile_back.gif" height="32">&nbsp;<img height="8" src="images/nav_m.gif" width="8" border="0">
			<b><font color="#FFFFFF">登陆</font></a></b></td>
		</tr>
		<tr>
			<td class="pformstrip">请在下方输入你的资料来进行登陆.</td>
		</tr>
		<tr>
			<td height="130" align="center">
			<table border="0" cellspacing="0" cellpadding="0" style="border: 1px solid #992A2A" width="99%">
				<tr>
					<td bgcolor="#E3C0C0" height="27"><b>&nbsp;<font color="#992A2A">Attention!
					</font></b></td>
				</tr>
				<tr>
					<td bgcolor="#F2DDDD" height="90">
					<p><font color="#992A2A">&nbsp;在你能够登陆之前,你必须已经注册一个帐号.<br>
&nbsp;如果你还没有帐号, 你可以点击画面顶端的&#39;注册&#39;连结进行注册.</font></p>
					<p><b><font color="#992A2A">&nbsp;我忘记了我的密码! </font>
					<a href="Reg.asp">
					<font color="#992A2A">按这里!</font></a></b> </p>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td height="100" align="center">
			<table cellspacing="1" width="99%">
				<tr>
					<td width="60%"><fieldset>
					<legend><b>登陆</b></legend>
					<table cellspacing="0" width="100%" height="70" cellpadding="0">
						<tr>
							<td width="50%"><b>&nbsp;请输入你的名称</b></td>
							<td width="50%">
							<input class="InputBox" type="text" name="Users_Name" style="height:18; width=140"></td>
						</tr>
						<tr>
							<td width="50%"><b>&nbsp;请输入你的密码</b></td>
							<td width="50%">
							<input class="InputBox" type="password" name="Users_Password" style="height:18; width=140"></td>
						</tr>
					</table>
					</fieldset> </td>
					<td width="40%"><fieldset>
					<legend><b>选项</b></legend>
					<table cellspacing="1" id="table3" width="100%" height="70">
						<tr>
							<td width="10%">
							<input type="checkbox" checked value="1" name="CookiesTime" class="InputBox"></td>
							<td width="90%"><b>记住我?</b><br>
							<span class="desc">不推荐在共享的计算机上使用</span></td>
						</tr>
						<tr>
							<td width="10%">
							<input type="checkbox" value="1" name="Privacy" class="InputBox" disabled></td>
							<td width="90%"><b>登陆不可用</b><br>
							<span class="desc">不要显示在在线列表</span></td>
						</tr>
					</table>
					</fieldset> </td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td class="pformstrip" align="center">
			<input class="InputBox" type="submit" value="登陆" name="submit"></td>
			</td>
		</tr>
	</form>
</table>
<!-- #include file="CopyRight.asp"-->