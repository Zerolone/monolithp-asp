<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><table bgcolor="#F0F5FA" cellspacing="1" width="98%" align="center" height="30" style="border: 1px solid #C2CFDF" id="table2">
	<%if IsLogin=1 then%>
	<form method="post" action="Search.asp" name="FrmSearch">
		<tr>
			<td>
			<table cellspacing="0" cellpadding="0" width="100%" height="100%">
				<tr>
					<td>&nbsp;<b>欢迎回来; 您最后访问是在: <%=Request.Cookies(Monolith_CookiesName)("Users_LastLogin")%></b></td>
					<td width="50%"  align="right">
					<input onfocus="this.value=''" type="text"  name="Search" value="输入关键字搜索..." class="InputBox" style="height:18; width=180">
					<input class="button" type="image" src="images/login-button.gif" name="I1" align="middle">
					</td>
				</tr>
			</table>
			</td>
		</tr>
	</form>
	<%else%>
	<form method="post" action="LoginCheck.asp" name="FrmLogin">
		<tr>
			<td>
			<table cellspacing="0" cellpadding="0" width="100%" height="100%">
				<tr>
					<td>&nbsp;<b>欢迎回来</b></td>
					<td width="50%" align="right">
					<input onfocus="this.value=''" type="text"  name="Users_Name" value="UserName" class="InputBox" style="height:18; width=130">
					<input onfocus="this.value=''" type="password" name="Users_PassWord" value="PassWord" class="InputBox" style="height:18; width=130">
					<input class="button" type="image" src="images/login-button.gif" name="I1" align="middle">
					</td>
				</tr>
			</table>
			</td>
		</tr>
	</form>
	<%end if%>
</table>