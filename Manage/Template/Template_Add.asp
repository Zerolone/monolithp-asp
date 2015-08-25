<!-- #include virtual ="/Inc/Monolith_Conn.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.css" rel="stylesheet" type="text/css">
<script language="javascript" src="../../Js/News_Edit.JS"></script>
<script language="javascript">
function check()
{
  var Check_Template_Parent_ID	= document.FrmAddTemplate.Template_Parent_ID;
  if (Check_Template_Parent_ID.value == "")
  {
 	alert("请输入上级ID！");
	Check_Template_Parent_ID.focus();
	return false;
  }
  
  var Check_Template_Title	= document.FrmAddTemplate.Template_Title;
  if (Check_Template_Title.value == "")
  {
 	alert("请输入模版标题！");
	Check_Template_Title.focus();
	return false;
  }
}

function fnPreHandle()
{
var iCount; //拆分为多少个域
var strData; //原始数据
var iMaxChars = 50000;//考虑到汉字为双字节，域的最大字符数限制为50K
var iBottleNeck = 2000000;//如果文章超过2M字，需要提示用户
var strHTML;

strData = FrmAddTemplate.Template_Content.value;

if (strData.length > iBottleNeck)
{
if (confirm("您要提交的文本太长，建议您拆分为几部分分别发布。\n如果您坚持提交，注意需要较长时间才能提交成功。\n\n是否坚持提交？") == false)
return false;
}
iCount = parseInt(strData.length / iMaxChars) + 1;
strHTML = "<input type=hidden name=hdnCount value=" + iCount + ">";

for (var i = 1; i <= iCount; i++)
{
	strHTML = strHTML + "\n" + "<input type=hidden name=hdnTemplate_Content" + i + ">";
}

document.all.divHidden.innerHTML = strHTML;

for (var i = 1; i <= iCount; i++)
{
FrmAddTemplate.elements["hdnTemplate_Content" + i].value = strData.substring((i - 1) * iMaxChars, i * iMaxChars);
}

FrmAddTemplate.Template_Content.value = "";

//alert(FrmAddTemplate.hdnCount.value);
}
</script>
<%
	Dim Template_Level_Code, Template_Parent_ID, Template_Level
	Template_Level_Code		= Request("Template_Level_Code")
	Template_Parent_ID		= Request("Template_Parent_ID")
	Template_Level				= Request("Template_Level")

	if Len(Template_Level) >= 50 then
		Response.write "不能再添加下级了！"
		Response.end
	end if

	Dim Template_Level_List, Template_Bound, Template_Level_Len
	Template_Level_List		= ""
	Template_Bound				= ""
	i											= 0

	Sql = "Select [Template_Level] From [Monolith_Template] Where [Template_Parent_ID]= " & Template_Parent_ID & " Order By [Template_Level]"
	Rs.Open Sql, Conn , 1, 1
	Do While Not (Rs.Bof Or Rs.Eof)
		if Template_Level_List = "" then
			Template_Level_Len	= Len(Rs(0))
			Template_Bound			= Left(Rs(0), Len(Rs(0))-2)
			Template_Level_List = Rs(0)
		else
			Template_Level_List = Template_Level_List & "," & Rs(0)
		end if
		i = i + 1
		if i mod 10  = 0 then Template_Level_List = Template_Level_List & "<br>" 
		Rs.MoveNext
	Loop
   Call CloseRs
	Call CloseConn
  if Len(Template_Bound) mod 2 = 1 then Template_Bound = "0" & Template_Bound
  if Template_Bound = "0" then Template_Bound = ""
		if Template_Level_List = "" then
				Template_Bound	= Template_Level
				Template_Level_Len	= Len(Template_Level) + 2
		end if
%>
<form method="POST" action="Template_Add_Submit.asp"  onSubmit="fnPreHandle();return check();" name="FrmAddTemplate">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;模版管理 &gt;&gt; 模版添加</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">模版ID：</font></td>
			<td bgcolor="#FFFFFF">系统自动生成</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">父模版ID：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Template_Parent_ID" size="40" value="<%=Template_Parent_ID%>" style="width=250px;" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（如果存在数字，则可以不用修改）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">模版排序等级：</font></td>
			<td bgcolor="#FFFFFF"><select size="1" name="Template_Level" style="width=250px;" class="InputBox"><%
				Dim TheString, i
				For i=1 to 99
					if Len(i) = 1 then
						TheString = Template_Bound & "0" & i
					else
						TheString = Template_Bound & i
					end if
					if instr(Template_Level_List, TheString) = 0 then
				%>
			<option value="<%=TheString%>"><%=TheString%></option>
			<%
					end if
				Next
				%></select> <font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（长度为[<font color="#FF0000"><b><%=Template_Level_Len%></b></font>]，可用范围[<%=Template_Bound%>01-<%=Template_Bound%>99]，已用<br><%=Template_Level_List%>，尽量不要使用重复）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">模版标题：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Template_Title" size="40" style="width=250px;" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（请在8个汉字或16个英文字母范围内）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">是否存在子模版：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="radio" value="true" name="Template_HasChild" id="Template_HasChild_1">是，做为子栏目<input type="radio" value="false" checked name="Template_HasChild" id="Template_HasChild_0">否，作为叶子</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（如作为节点请选择是）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">模版内容：</font></td>
			<td bgcolor="#FFFFFF">
			<textarea rows="8" name="Template_Content" id="Template_Content" cols="100"></textarea></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（如作为最子栏目则不需要输入模版内容）</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF"><input type="hidden" value="Same_Level" name="Template_Level_Code"><input type="submit" value=" 提 交 " name="B1" class="InputBox">
			 <input type="button" value=" 预 览 模 版 " name="B2" class="InputBox" onclick="runCode(Template_Content)">
			<input type="button" value=" 取 消 返 回 " name="B3" onclick="window.open('Template_Setting.asp','_self');" class="InputBox"><div id=divHidden></div></td>
		</tr>
	</table>
</form>
<%
'	end if
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->