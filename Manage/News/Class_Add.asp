<!-- #include virtual ="/Inc/Monolith_Conn.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript">
function check()
{
  var Check_News_Class_Parent_ID	= document.FrmAddMenu.News_Class_Parent_ID;
  if (Check_News_Class_Parent_ID.value == "")
  {
 	alert("请输入上级栏目ID！");
	Check_News_Class_Parent_ID.focus();
	return false;
  }
  
  var Check_News_Class_Title	= document.FrmAddMenu.News_Class_Title;
  if (Check_News_Class_Title.value == "")
  {
 	alert("请输入菜单标题！");
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
		Response.write "不能再添加下级了！"
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
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;新闻栏目管理 &gt;&gt; 栏目添加</font></b></td>
		</tr>		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">栏&nbsp;&nbsp; 目&nbsp;&nbsp;&nbsp;&nbsp; ID：</font></td>
			<td bgcolor="#FFFFFF">系统自动生成</td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">父&nbsp;&nbsp; 栏&nbsp;&nbsp; 目ID：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="News_Class_Parent_ID" size="40" value="<%=News_Class_Parent_ID%>" style="width=250px;" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">　</td>
			<td bgcolor="#FFFFFF">（如果存在数字，则可以不用修改）</td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">栏 目 排序等级：</font></td>
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
			<td align="right" width="100" height="20" bgcolor="#999999">　</td>
			<td bgcolor="#FFFFFF">（长度为[<font color="#FF0000"><b><%=News_Class_Level_Len%></b></font>]，可用范围[<%=News_Class_Bound%>01-<%=News_Class_Bound%>99]，已用[<%=News_Class_Level_List%>]，尽量不要使用重复）</td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">栏&nbsp; 目&nbsp; 标&nbsp; 题：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="News_Class_Title" size="40" style="width=250px;" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">　</td>
			<td bgcolor="#FFFFFF">（请在8个汉字或16个英文字母范围内）</td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">
			<font color="#FFFFFF">是否存在子栏目：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="radio" value="true" name="News_Class_HasChild" id="News_Class_HasChild_1"><label for="News_Class_HasChild_1">是，做为子栏目</label><input type="radio" value="false" checked name="News_Class_HasChild" id="News_Class_HasChild_0"><label for="News_Class_HasChild_0">否，作为叶子</label></td>
		</tr>
		<tr>
			<td align="right" width="100" height="20" bgcolor="#999999">　</td>
			<td bgcolor="#FFFFFF">（如作为节点请选择是）</td>
		</tr>
		<tr>
			<td width="100" height="20" bgcolor="#999999">　</td>
			<td bgcolor="#FFFFFF"><input type="hidden" value="Same_Level" name="News_Class_Level_Code"><input type="submit" value=" 提 交 " name="B1" class="InputBox">
			<input type="reset" value=" 重 置 " name="B2" class="InputBox"></td>
		</tr>
	</table>
</form>
<%
'	end if
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->