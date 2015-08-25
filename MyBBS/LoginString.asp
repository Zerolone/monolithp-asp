<!-- #include virtual = "/Inc/LoginUserCheck.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<%
	if IsLogin = 1 then
%>
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><table bgcolor="#F0F5FA" cellspacing="1" width="98%" align="center" height="26" style="border: 1px solid #C2CFDF">
	<tr>
		<td>
		<table cellspacing="0" cellpadding="0" width="100%">
			<tr>
				<td align="right"><b>
				<a title="编辑个人选项, 例如个人签名, 头像等等..." href="#">我的控制台</a></b> ·
				<a href="http://www.11k.net/bbs/index.php?act=Search&CODE=getnew">
				查看最新文章</a> ·
				<a title="浏览新话题, 管理团队等等..." href="javascript:buddy_pop();">我的助手</a> 
				· <a href="http://www.11k.net/bbs/index.php?act=Msg&CODE=01">0 个新短消息</a>
				</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<%else%>
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><table bgcolor="#F4E7EA" cellspacing="1" width="98%" align="center" height="26" style="border: 1px solid #986265">
	<tr>
		<td align="center"><b>欢迎你客人</b> ( <a href="Login.asp">登录</a> | <a href="Reg.asp">注册</a> 
		)</td>
	</tr>
</table>
<%end if%>