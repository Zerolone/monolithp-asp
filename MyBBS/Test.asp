
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<title>MyBBS (Power By Monolith)</title>
<style>
<!--
td
	{
		font-family: Tahoma,Verdana, Tahoma, Arial,Helvetica, sans-serif; 
		font-size: 12px; 
		word-break: break-all 
	}

body

	{
		font-family: Tahoma,Verdana, Tahoma, Arial,Helvetica, sans-serif; 
		font-size: 12px; 
		word-break: break-all 
	}

a:link, a:visited, a:active
	{	
		background: transparent;	
		color: #3A4F6C;	
		text-decoration: none;
	}
	
a:hover
	{
		color: #5176B5;
	}

.InputBox
	{	
		border: 1px solid #999999
	}

.desc { color:#434951 }

.forminput  { font-size: 11px; font-family: verdana, helvetica, sans-serif; vertical-align: middle }

.Board_Top
	{
		background-image: url("/images/tile_back.gif");
		height:68px;
		font-face:impact;
		font-size:36;
		color:#FFFFFF;
		font-weight:bold;
	}
	
.Board_Table
	{
		border: 1px solid #000000;
	}

.Board_Title
	{
		background-image: url("images/tile_back.gif");
		height:32; 
		padding:6; 
		margin-top:1;
		border-top:1;
		color:#FFFFFF;  
		font-weight:bold
	}

.Board_Sub
	{
		background-image: url("images/tile_sub.gif"); 
		height:24; 
		padding:6; 
		margin-top:1;
		border-top:1;  
		color: #3A4F6C;
		font-weight:bold
	}
	
.Board_Foot
	{
		background-color: #BCD0ED; 
		color:#3A4F6C;
	}

.List_Split
	{
		background-color: #C2CFDF; 
		color:#4C77B6;
		height:24; 
	}	

.List_Head
	{ 
		background-color: #D1DCEB; 
		color:#3A4F6C;
		font-weight:bold;
		padding:7px;
		margin-top:1px 
	}
	
.List_Left
	{ 
		background-color: #F5F9FD; 
		padding:6px; 
		margin-top:1px;
		border-top:1px solid #C2CFDF; 
		border-right:1px solid #C2CFDF; 
	}

.List_Right
	{
		background-color: #F5F9FD; 
		padding:6px; 
		margin-top:1px;
		border-top:1px solid #C2CFDF; 
	}

.row1 { background-color: #E4EAF2 }
.row2 { background-color: #DFE6EF }	
-->
</style>

</head>

<body background="/images/bg.GIF">

<table width="98%" align="center" cellspacing="0" border="0" class="Board_Table">
	<tr>
		<td class="Board_Top">&nbsp;Monolith MyBBS	<font size="3">Beta V1.0</font></td>
	</tr>
	<tr height="26">
		<td class="Board_Sub">
		<table cellspacing="0" cellpadding="0" width="100%">
			<tr>
				<td><a href="Default.asp"> Welcome to MyBBs.   Current 
				Version 1.0(������)</a></td>
				<td align="right"><font color="#3A4F6C">��<a href="/Default.asp">��ҳ</a>��<a href="../Bulletin/Default.asp">����</a>��<a href="Default.asp">��̳</a>��<a href="#">����</a>��<a href="#">�û�</a>��<a href="#">����</a>��<a target="_blank" href="/Help/Default.asp">ʹ�ð���</a>��</font>&nbsp;&nbsp;&nbsp;
				</td>
			</tr>
		</table>
		</td>
	</tr>
</table>

<!--�ָ���--><img src="images/Blank.gif" border="0" width="1" height="5"><!--�ָ���--><table bgcolor="#F4E7EA" cellspacing="1" width="98%" align="center" height="26" style="border: 1px solid #986265">
	<tr>
		<td align="center"><b>��ӭ�����</b> ( <a href="Login.asp">��¼</a> | <a href="Reg.asp">ע��</a> 
		)</td>
	</tr>
</table>

<!--�ָ���--><img src="images/Blank.gif" border="0" width="1" height="5"><!--�ָ���--><table cellspacing="0" cellpadding="0" width="98%" align="center">
	<tr>
		<td><b><a href="/MyBBS/Default.asp"><img src="/images/nav.gif" border="0"> MyBBS Board</a> 
		
		</b>
		</td>
	</tr>
</table><!--�ָ���--><img src="images/Blank.gif" border="0" width="1" height="5"><!--�ָ���--><table bgcolor="#F0F5FA" cellspacing="1" width="98%" align="center" height="30" style="border: 1px solid #C2CFDF" id="table2">
	
	<form method="post" action="LoginCheck.asp" name="FrmLogin">
		<tr>
			<td>
			<table cellspacing="0" cellpadding="0" width="100%" height="100%">
				<tr>
					<td>&nbsp;<b>��ӭ����</b></td>
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
	
</table>
<!--�ָ���--><img src="images/Blank.gif" border="0" width="1" height="5"><!--�ָ���-->
<table width="98%" align="center" cellspacing="0" cellpadding="0" class="Board_Table">
	<tr>
		<td class="Board_Title">
		<img height="8" src="images/nav_m.gif" width="8" border="0"> <a href="#">
		<font color="#FFFFFF">��̳������</font></a></td>
	</tr>
	<tr>
		<td>
		<table cellspacing="1" cellpadding="4" width="100%" border="0">
			<tr>
				<td align="middle" width="2%" class="Board_Sub">��</td>
				<td width="59%" class="Board_Sub">���</td>
				<td align="middle" width="7%" class="Board_Sub">����</td>
				<td align="middle" width="7%" class="Board_Sub">�ظ�</td>
				<td width="25%" class="Board_Sub">���·�����Ϣ</td>
			</tr>
			
			<tr>
				<td class="row1" align="middle">
				<img alt="û���µ�����" src="images/bf_nonew.gif" border="0"></td>
				<td class="row1"><b>
				<a href="List.asp?Board_ID=2">
				<u><font color="#000000">��̳����</font></u></a></b><br>
				<font color="#434951">�Ǻ�������ǲ�����̳�Ķ���<br>
				<i>����: <a href="ViewUser.asp?Users_ID=11">Zerolone</a></i></font></td>
				<td class="row2" align="middle">34</td>
				<td class="row2" align="middle">30</td>
				<td class="row2" nowrap>2005-3-29 23:48:34<br>
				<b>����:</b>
				<a title="ǰ����һƪδ��������" href="ListView.asp?BBS_ID=143&Board_ID=2">
				������Ҫ</a><br>
				<b>����:</b>
				<a href="ViewUser.asp?Users_ID=11">Zerolone</a></td>
			</tr>
			
			<tr>
				<td class="row1" align="middle">
				<img alt="û���µ�����" src="images/bf_nonew.gif" border="0"></td>
				<td class="row1"><b>
				<a href="List.asp?Board_ID=5">
				<u><font color="#000000">����</font></u></a></b><br>
				<font color="#434951">�����йظ���̳�Ĺ��涼�����ڴ�<br>
				<i>����: <a href="ViewUser.asp?Users_ID=0">��</a></i></font></td>
				<td class="row2" align="middle">3</td>
				<td class="row2" align="middle">2</td>
				<td class="row2" nowrap>2004-6-13 19:50:22<br>
				<b>����:</b>
				<a title="ǰ����һƪδ��������" href="ListView.asp?BBS_ID=0&Board_ID=5">
				�Ǻ�,����</a><br>
				<b>����:</b>
				<a href="ViewUser.asp?Users_ID=0">Zerolone</a></td>
			</tr>
			
		</table>
		</td>
	</tr>
	<tr>
		<td class="Board_Foot">��</td>
	</tr>
</table>
<br><table width="98%" align="center" cellspacing="0" cellpadding="0" class="Board_Table">
	<tr>
		<td class="Board_Title">
		<img height="8" src="images/nav_m.gif" width="8" border="0"> <a href="#">
		<font color="#FFFFFF">��̳ˮ��</font></a></td>
	</tr>
	<tr>
		<td>
		<table cellspacing="1" cellpadding="4" width="100%" border="0">
			<tr>
				<td align="middle" width="2%" class="Board_Sub">��</td>
				<td width="59%" class="Board_Sub">���</td>
				<td align="middle" width="7%" class="Board_Sub">����</td>
				<td align="middle" width="7%" class="Board_Sub">�ظ�</td>
				<td width="25%" class="Board_Sub">���·�����Ϣ</td>
			</tr>
			
			<tr>
				<td class="row1" align="middle">
				<img alt="û���µ�����" src="images/bf_nonew.gif" border="0"></td>
				<td class="row1"><b>
				<a href="List.asp?Board_ID=6">
				<u><font color="#000000">FFF</font></u></a></b><br>
				<font color="#434951">��<br>
				<i>����: <a href="ViewUser.asp?Users_ID=3">��</a></i></font></td>
				<td class="row2" align="middle">1</td>
				<td class="row2" align="middle">0</td>
				<td class="row2" nowrap>2004-12-25 21:17:16<br>
				<b>����:</b>
				<a title="ǰ����һƪδ��������" href="ListView.asp?BBS_ID=0&Board_ID=6">
				�Ǻ�  �� �������</a><br>
				<b>����:</b>
				<a href="ViewUser.asp?Users_ID=3">Zerolone</a></td>
			</tr>
			
			<tr>
				<td class="row1" align="middle">
				<img alt="û���µ�����" src="images/bf_nonew.gif" border="0"></td>
				<td class="row1"><b>
				<a href="List.asp?Board_ID=8">
				<u><font color="#000000">HHH</font></u></a></b><br>
				<font color="#434951">��<br>
				<i>����: <a href="ViewUser.asp?Users_ID=0">��</a></i></font></td>
				<td class="row2" align="middle">0</td>
				<td class="row2" align="middle">0</td>
				<td class="row2" nowrap>2004-6-17 0:08:15<br>
				<b>����:</b>
				<a title="ǰ����һƪδ��������" href="ListView.asp?BBS_ID=0&Board_ID=8">
				��</a><br>
				<b>����:</b>
				<a href="ViewUser.asp?Users_ID=0">��</a></td>
			</tr>
			
		</table>
		</td>
	</tr>
	<tr>
		<td class="Board_Foot">��</td>
	</tr>
</table>
<br>
<table width="98%" align="center" cellspacing="0" cellpadding="0" class="Board_Table">
	<tr>
		<td class="Board_Title">&nbsp;<img height="8" src="images/nav_m.gif" width="8" border="0"> 
		��̳ͳ��</td>
	</tr>
	<tr>
		<td>
		<table cellspacing="1" cellpadding="4" width="100%" border="0">
			<tr>
				<td colspan="2" class="List_Head">��̳ͳ��</td>
			</tr>
			<tr>
				<td class="row2" width="5%">
				<img alt="��̳����ͳ��" src="images/stats.gif" border="0" width="30" height="30"></td>
				<td class="row1" align="left" width="95%">���ǵĻ�Ա�ܹ������� <b>200</b> 
				ƪ����<br>
				������ <b>10</b> λע���Ա<br>
				��ӭ���¼���Ļ�Ա: <b><a href="ViewUser.asp?Users_ID=12">
				qqq</a></b><br>
				��߼�¼������ <b>2004-12-26</b> �� <b>8</b> λʹ����ͬʱ������.
				</td>
			</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td class="Board_Foot">��</td>
	</tr>
</table>
<!--�ָ���--><img src="images/Blank.gif" border="0" width="1" height="5"><!--�ָ���-->
<table bgcolor="#8394B2" cellspacing="0" width="98%" align="center" height="26" cellpadding="0">
	<tr>
		<td>
		<table cellspacing="0" cellpadding="0" width="100%">
			<tr>
				<td align="right" width="50%"><a href="#"><b>
				<font color="#FFFFFF">�򻯰汾</font></b></a></td>
				<td align="right" width="50%">������ʱ��: 2005-4-21 21:29:03&nbsp;&nbsp;</td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<!--�ָ���--><img src="images/Blank.gif" border="0" width="1" height="5"><!--�ָ���--><table bgcolor="#EEEEEE" cellspacing="0" width="98%" align="center" height="26" cellpadding="0">
	<tr>
		<td align="center">�� Powered by Monolith - MyBBs<br>[ Script Execution time: 2,015.62500���롣 ]   [ 5 queries used ] </td>
	</tr>
</table>