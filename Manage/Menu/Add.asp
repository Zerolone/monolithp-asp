<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript">
function check()
{
  var Check_Tree_Parent_ID	= document.FrmAddMenu.Tree_Parent_ID;
  if (Check_Tree_Parent_ID.value == "")
  {
 	alert("请输入父菜单ID！");
	Check_Tree_Parent_ID.focus();
	return false;
  }
  
  var Check_Tree_Title	= document.FrmAddMenu.Tree_Title;
  if (Check_Tree_Title.value == "")
  {
 	alert("请输入菜单标题！");
	Check_Tree_Title.focus();
	return false;
  }
}
</script>

<%
	Dim Tree_Level_Code, Tree_Parent_ID, Tree_Level
	Tree_Level_Code		= Request("Tree_Level_Code")
	Tree_Parent_ID		= Request("Tree_Parent_ID")
	Tree_Level				= Request("Tree_Level")

	if Len(Tree_Level) >= 50 then
		Response.write "不能再添加下级了！"
		Response.end
	end if

	Dim Tree_Level_List, Tree_Bound, Tree_Level_Len
	Tree_Level_List		= ""
	Tree_Bound				= ""

'	if Tree_Level_Code = "Same_Level" then
		Sql = "Select [Tree_Level] From [Monolith_Tree] Where [Tree_Parent_ID]= " & Tree_Parent_ID & " Order By Tree_Level"
		Rs.Open Sql, Conn , 1, 1
		Do While Not (Rs.Bof Or Rs.Eof)
			if Tree_Level_List = "" then
				Tree_Level_Len	= Len(Rs(0))
'				Tree_Bound			= Rs(0) \ 100
				Tree_Bound			= Left(Rs(0), Len(Rs(0))-2)
				Tree_Level_List = Rs(0)

'---------------------------------------------------------------------------------------------------------
'				Tree_Level_List = Left(Rs(0), Tree_Level_Len-2) & "<font color=""#FF0000"">" & Right(Rs(0),2) & "</font>"
'---------------------------------------------------------------------------------------------------------
			else
				Tree_Level_List = Tree_Level_List & "," & Rs(0)

'---------------------------------------------------------------------------------------------------------
'				Tree_Level_List = Tree_Level_List & "," & Left(Rs(0), Tree_Level_Len-2) & "<font color=""#FF0000"">" & Right(Rs(0),2) & "</font>"
'---------------------------------------------------------------------------------------------------------
			end if
			Rs.MoveNext
		Loop
    Call CloseRs
		Call CloseConn
    if Len(Tree_Bound) mod 2 = 1 then Tree_Bound = "0" & Tree_Bound
    if Tree_Bound = "0" then Tree_Bound = ""
		if Tree_Level_List = "" then
				Tree_Bound	= Tree_Level
				Tree_Level_Len	= Len(Tree_Level) + 2
		end if
%>
<form method="POST" action="Add_Submit.asp"  onSubmit="return check()" name="FrmAddMenu">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;左侧菜单管理 &gt;&gt; 菜单添加</font></b></td>
		</tr>		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">菜&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 单&nbsp; ID：</font></td>
			<td bgcolor="#FFFFFF">系统自动生成</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">父&nbsp; 菜&nbsp; 单&nbsp; ID：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Tree_Parent_ID" size="40" value="<%=Tree_Parent_ID%>" style="width=250px;" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（如果存在数字，则可以不用修改）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">菜 单 排序等级：</font></td>
			<td bgcolor="#FFFFFF"><select size="1" name="Tree_Level" style="width=250px;" class="InputBox"><%
				Dim TheString, i
				For i=1 to 99
					if Len(i) = 1 then
						TheString = Tree_Bound & "0" & i
					else
						TheString = Tree_Bound & i
					end if
					if instr(Tree_Level_List, TheString) = 0 then
				%>
			<option value="<%=TheString%>"><%=TheString%></option>
			<%
					end if
				Next
				%></select> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（长度为[<font color="#FF0000"><b><%=Tree_Level_Len%></b></font>]，可用范围[<%=Tree_Bound%>01-<%=Tree_Bound%>99]，已用[<%=Tree_Level_List%>]，尽量不要使用重复）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">菜&nbsp; 单&nbsp; 标&nbsp; 题：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Tree_Title" size="40" style="width=250px;" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（请在8个汉字或16个英文字母范围内）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">菜&nbsp; 单&nbsp; 链&nbsp; 接：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Tree_Url" size="40" style="width=250px;" class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（如作为节点可以不填写）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">菜 单 链接框架：</font></td>
			<td bgcolor="#FFFFFF"><select size="1" name="Tree_Target" style="width=250px;" class="InputBox">
			<option value="tabWin">默认窗口</option>
			<option value="_self">相同框架</option>
			<option value="_top">整页</option>
			<option value="_blank">新窗口</option>
			<option value="_parent">父窗口</option>
			</select></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（如作为节点请选择为默认窗口）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">是否存在子菜单：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="radio" value="true" name="Tree_HasChild" id="Tree_HasChild_1"><label for="Tree_HasChild_1">是，做为子菜单</label><input type="radio" value="false" checked name="Tree_HasChild" id="Tree_HasChild_0"><label for="Tree_HasChild_0">否，作为叶子</label></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（如作为节点请选择是）</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF"><input type="hidden" value="Same_Level" name="Tree_Level_Code"><input type="submit" value=" 提 交 (Alt +Y) " name="B1" class="InputBox" accesskey="Y">
			<input type="reset" value=" 重 置 (Alt + N) " name="B2" class="InputBox" accesskey="N"></td>
		</tr>
	</table>
</form>
<table><tr><td height="160"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->