<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp"--><!-- #include virtual ="/Manage/Check.asp"--><%
	Dim ThisPage, Page
	Dim Bulletin_ID
	
	Page							= Request("Page")

  ThisPage					= Request.ServerVariables("PATH_INFO")
	if Request("Code") = "Edit" then
		Bulletin_ID = Request("Bulletin_ID")
		CheckNum Bulletin_ID
		
		Dim Bulletin_Title, Bulletin_Title2, Bulletin_Content, Bulletin_Content2

		Bulletin_Title			= Request("Bulletin_Title")
		Bulletin_Title2			= Request("Bulletin_Title2")
		Bulletin_Content		= Request("Bulletin_Content")
		Bulletin_Content2		= Request("Bulletin_Content2")

		Bulletin_Title			= MonolithEncode(Bulletin_Title)
		Bulletin_Title2			= MonolithEncode(Bulletin_Title2)
		Bulletin_Content		= MonolithEncode(Bulletin_Content)
		Bulletin_Content2		= MonolithEncode(Bulletin_Content2)
		
		'---------------------0------------------1-----------------2--------------------3
		Sql = "Select [Bulletin_Title], [Bulletin_Title2], [Bulletin_Content], [Bulletin_Content2] From [Monolith_Bulletin] Where [Bulletin_ID]=" & Bulletin_ID
		Rs.Open Sql, Conn, 1, 3
		Rs(0)								= Bulletin_Title
		Rs(1)								= Bulletin_Title2
		Rs(2)								= Bulletin_Content
		Rs(3)								= Bulletin_Content2
		Rs.Update
		Rs.Close
		
		Response.Redirect("Default.asp?Page=" & Page)

	end if
%>
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
<%
	Dim TrBgColor
	
	Bulletin_ID = Request("Bulletin_ID")
	CheckNum Bulletin_ID
	
	'----------------------0-----------------1------------------2-------------------3
	Sql = "Select [Bulletin_Title], [Bulletin_Title2], [Bulletin_Content], [Bulletin_Content2] From [Monolith_Bulletin] Where [Bulletin_ID]=" & Bulletin_ID
	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof or Rs.Eof then
	  Response.write "�����治���ڣ�"
		Response.end
	end if	
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<form method="POST" action="Edit.asp" onsubmit="return check()" name="FrmAddMenu">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;������� &gt;&gt; �����޸�</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� ��&nbsp;&nbsp;&nbsp;&nbsp; ID��</font></td>
			<td bgcolor="#FFFFFF"><%=Bulletin_ID%>��ϵͳ�Զ����ɣ�</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� ��&nbsp; �� �⣺</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Bulletin_Title" size="40" style="width=250px;" value="<%=Rs(0)%>" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� �� �����⣺</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Bulletin_Title2" size="40" style="width=250px;" value="<%=Rs(1)%>" class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�����������ݣ�</font></td>
			<td bgcolor="#FFFFFF"><font color="#FF0000">
			<textarea rows="2" name="Bulletin_Content2" cols="20" style="width=500;height:100" class="InputBox"><%=Rs(3)%></textarea></font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">��һ��Ϊ�������ݣ�</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�� ��&nbsp; �� �ݣ�</font></td>
			<td bgcolor="#FFFFFF"><font color="#FF0000">
			<textarea rows="2" name="Bulletin_Content" cols="20" style="width=500;height:100" class="InputBox"><%=Rs(2)%></textarea> 
			*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">��������ʾ���ݣ�</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF"><input type="hidden" value="Edit" name="Code">
			<input type="hidden" value="<%=Bulletin_ID%>" name="Bulletin_ID">
			<input type="hidden" value="<%=Page%>" name="Page">
			<input type="submit" value=" �� �� (Alt + S) " name="B1" class="InputBox" accesskey="A">
			<input type="reset" value=" �� �� (Alt + N) " name="B2" class="InputBox" accesskey="N" onclick="window.open('Default.asp?Page=<%=Page%>','_self');"></td>
		</tr>
	</form>
</table>
<table><tr><td height="144"></td></tr></table>
<%
	Call CloseRs
	Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->