<!-- #include virtual = "/Inc/LoginUserCheck.asp" -->
<!-- #include virtual = "/Inc/Monolith_Function.asp" -->
<%
	if IsLogin = 1 then
%>
<!--�ָ���--><img src="images/Blank.gif" border="0" width="1" height="5"><!--�ָ���--><table bgcolor="#F0F5FA" cellspacing="1" width="98%" align="center" height="26" style="border: 1px solid #C2CFDF">
	<tr>
		<td>
		<table cellspacing="0" cellpadding="0" width="100%">
			<tr>
				<td align="right"><b>
				<a title="�༭����ѡ��, �������ǩ��, ͷ��ȵ�..." href="#">�ҵĿ���̨</a></b> ��
				<a href="http://www.11k.net/bbs/index.php?act=Search&CODE=getnew">
				�鿴��������</a> ��
				<a title="����»���, �����Ŷӵȵ�..." href="javascript:buddy_pop();">�ҵ�����</a> 
				�� <a href="http://www.11k.net/bbs/index.php?act=Msg&CODE=01">0 ���¶���Ϣ</a>
				</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<%else%>
<!--�ָ���--><img src="images/Blank.gif" border="0" width="1" height="5"><!--�ָ���--><table bgcolor="#F4E7EA" cellspacing="1" width="98%" align="center" height="26" style="border: 1px solid #986265">
	<tr>
		<td align="center"><b>��ӭ�����</b> ( <a href="Login.asp">��¼</a> | <a href="Reg.asp">ע��</a> 
		)</td>
	</tr>
</table>
<%end if%>