<!-- #include virtual ="/Inc/Monolith_Conn.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript">
function check()
{
  var Check_News_Class_Parent_ID	= document.FrmAddMenu.News_Class_Parent_ID;
  if (Check_News_Class_Parent_ID.value == "")
  {
 	alert("�������ϼ���ĿID��");
	Check_News_Class_Parent_ID.focus();
	return false;
  }
  
  var Check_News_Class_Title	= document.FrmAddMenu.News_Class_Title;
  if (Check_News_Class_Title.value == "")
  {
 	alert("������˵����⣡");
	Check_News_Class_Title.focus();
	return false;
  }
}
</script>

<%
	Dim News_Class_Level_Code, News_Class_Parent_ID, News_Class_Level
	News_Class_Level_Code		= Request("News_Class_Level_Code")
	News_Class_Parent_ID		= Request("News_Class_Parent_ID")
	News_Class_Level				= Request("News_Class_Level")

	if Len(News_Class_Level) >= 50 then
		Response.write "����������¼��ˣ�"
		Response.end
	end if

	Dim News_Class_Level_List, News_Class_Bound, News_Class_Level_Len
	News_Class_Level_List		= ""
	News_Class_Bound				= ""

	Sql = "Select [News_Class_Level] From [Monolith_News_Class] Where [News_Class_Parent_ID]= " & News_Class_Parent_ID & " Order By [News_Class_Level]"
	Rs.Open Sql, Conn , 1, 1
	Do While Not (Rs.Bof Or Rs.Eof)
		if News_Class_Level_List = "" then
			News_Class_Level_Len	= Len(Rs(0))
			News_Class_Bound			= Left(Rs(0), Len(Rs(0))-2)
			News_Class_Level_List = Rs(0)
		else
			News_Class_Level_List = News_Class_Level_List & "," & Rs(0)
		end if
		Rs.MoveNext
	Loop
   Call CloseRs
	Call CloseConn
  if Len(News_Class_Bound) mod 2 = 1 then News_Class_Bound = "0" & News_Class_Bound
  if News_Class_Bound = "0" then News_Class_Bound = ""
		if News_Class_Level_List = "" then
				News_Class_Bound	= News_Class_Level
				News_Class_Level_Len	= Len(News_Class_Level) + 2
		end if
%>
<form method="POST" action="Class_Add_Submit.asp"  onSubmit="return check()" name="FrmAddMenu">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;������Ŀ���� &gt;&gt; ��Ŀ���</font></b></td>
		</tr>		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">��&nbsp;&nbsp; Ŀ&nbsp;&nbsp;&nbsp;&nbsp; ID��</font></td>
			<td bgcolor="#FFFFFF">ϵͳ�Զ�����</td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">��&nbsp;&nbsp; ��&nbsp;&nbsp; ĿID��</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="News_Class_Parent_ID" size="40" value="<%=News_Class_Parent_ID%>" style="width=250px;" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">��</td>
			<td bgcolor="#FFFFFF">������������֣�����Բ����޸ģ�</td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">�� Ŀ ����ȼ���</font></td>
			<td bgcolor="#FFFFFF"><select size="1" name="News_Class_Level" style="width=250px;" class="InputBox"><%
				Dim TheString, i
				For i=1 to 99
					if Len(i) = 1 then
						TheString = News_Class_Bound & "0" & i
					else
						TheString = News_Class_Bound & i
					end if
					if instr(News_Class_Level_List, TheString) = 0 then
				%>
			<option value="<%=TheString%>"><%=TheString%></option>
			<%
					end if
				Next
				%></select> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">��</td>
			<td bgcolor="#FFFFFF">������Ϊ[<font color="#FF0000"><b><%=News_Class_Level_Len%></b></font>]�����÷�Χ[<%=News_Class_Bound%>01-<%=News_Class_Bound%>99]������[<%=News_Class_Level_List%>]��������Ҫʹ���ظ���</td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">��&nbsp; Ŀ&nbsp; ��&nbsp; �⣺</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="News_Class_Title" size="40" style="width=250px;" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">��</td>
			<td bgcolor="#FFFFFF">������8�����ֻ�16��Ӣ����ĸ��Χ�ڣ�</td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">�Ƿ��������Ŀ��</font></td>
			<td bgcolor="#FFFFFF">
			<input type="radio" value="true" name="News_Class_HasChild" id="News_Class_HasChild_1"><label for="News_Class_HasChild_1">�ǣ���Ϊ����Ŀ</label><input type="radio" value="false" checked name="News_Class_HasChild" id="News_Class_HasChild_0"><label for="News_Class_HasChild_0">����ΪҶ��</label></td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">��</td>
			<td bgcolor="#FFFFFF">������Ϊ�ڵ���ѡ���ǣ�</td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999">��</td>
			<td bgcolor="#FFFFFF"><input type="hidden" value="Same_Level" name="News_Class_Level_Code"><input type="submit" value=" �� �� " name="B1" class="InputBox">
			<input type="reset" value=" �� �� " name="B2" class="InputBox"></td>
		</tr>
	</table>
</form>
<%
'	end if
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->