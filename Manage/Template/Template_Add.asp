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
 	alert("�������ϼ�ID��");
	Check_Template_Parent_ID.focus();
	return false;
  }
  
  var Check_Template_Title	= document.FrmAddTemplate.Template_Title;
  if (Check_Template_Title.value == "")
  {
 	alert("������ģ����⣡");
	Check_Template_Title.focus();
	return false;
  }
}

function fnPreHandle()
{
var iCount; //���Ϊ���ٸ���
var strData; //ԭʼ����
var iMaxChars = 50000;//���ǵ�����Ϊ˫�ֽڣ��������ַ�������Ϊ50K
var iBottleNeck = 2000000;//������³���2M�֣���Ҫ��ʾ�û�
var strHTML;

strData = FrmAddTemplate.Template_Content.value;

if (strData.length > iBottleNeck)
{
if (confirm("��Ҫ�ύ���ı�̫�������������Ϊ�����ֱַ𷢲���\n���������ύ��ע����Ҫ�ϳ�ʱ������ύ�ɹ���\n\n�Ƿ����ύ��") == false)
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
		Response.write "����������¼��ˣ�"
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
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;ģ����� &gt;&gt; ģ�����</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">ģ��ID��</font></td>
			<td bgcolor="#FFFFFF">ϵͳ�Զ�����</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">��ģ��ID��</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Template_Parent_ID" size="40" value="<%=Template_Parent_ID%>" style="width=250px;" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">������������֣�����Բ����޸ģ�</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">ģ������ȼ���</font></td>
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
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">������Ϊ[<font color="#FF0000"><b><%=Template_Level_Len%></b></font>]�����÷�Χ[<%=Template_Bound%>01-<%=Template_Bound%>99]������<br><%=Template_Level_List%>��������Ҫʹ���ظ���</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">ģ����⣺</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Template_Title" size="40" style="width=250px;" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">������8�����ֻ�16��Ӣ����ĸ��Χ�ڣ�</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�Ƿ������ģ�棺</font></td>
			<td bgcolor="#FFFFFF">
			<input type="radio" value="true" name="Template_HasChild" id="Template_HasChild_1">�ǣ���Ϊ����Ŀ<input type="radio" value="false" checked name="Template_HasChild" id="Template_HasChild_0">����ΪҶ��</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">������Ϊ�ڵ���ѡ���ǣ�</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">ģ�����ݣ�</font></td>
			<td bgcolor="#FFFFFF">
			<textarea rows="8" name="Template_Content" id="Template_Content" cols="100"></textarea></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">������Ϊ������Ŀ����Ҫ����ģ�����ݣ�</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF"><input type="hidden" value="Same_Level" name="Template_Level_Code"><input type="submit" value=" �� �� " name="B1" class="InputBox">
			 <input type="button" value=" Ԥ �� ģ �� " name="B2" class="InputBox" onclick="runCode(Template_Content)">
			<input type="button" value=" ȡ �� �� �� " name="B3" onclick="window.open('Template_Setting.asp','_self');" class="InputBox"><div id=divHidden></div></td>
		</tr>
	</table>
</form>
<%
'	end if
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->