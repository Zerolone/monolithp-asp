<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript" src="/Js/News_Edit.JS"></script>
<script language="javascript">
function check()
{
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
	Dim RsTemp
	Set	RsTemp				= Server.CreateObject("ADODB.RecordSet")

	Dim ThisPage
	ThisPage			   = Request.ServerVariables("PATH_INFO")

	Dim Template_ID, Template_Parent_ID, Template_Parent_ID_Arr,Template_Parent_ID_Old, Template_Level, Template_Level_Len, Template_Level_Old, Template_Title, Template_HasChild

	Template_ID						= Request("Template_ID")
	if Request("Edit") = "True" then
		Template_Parent_ID		= Request("Template_Parent_ID")
		Template_Level				= Request("Template_Level")
		Template_Level_Old		= Request("Template_Level_Old")
		Template_Title				= Request("Template_Title")
		Template_HasChild			= Request("Template_HasChild")
		Template_Parent_ID_Old=	Request("Template_Parent_ID_Old")
		
		Template_Parent_ID_Arr	= Split(Template_Parent_ID, "|")
		
		'��ε�ʱ����
		if Template_Parent_ID_Arr(1) <> 0 then
			if Len(Template_Level) + 2 <> Len(Template_Parent_ID_Arr(1)) then
				Template_Level	= Template_Parent_ID_Arr(1) & Right(Template_Level, 2)
			end if
		else
				Template_Parent_ID_Arr(1)	= ""
				Template_Level	= Right(Template_Level, 2)
		end if

		Template_Title				= MonolithEncode(Template_Title)
		Template_Level_Len	= Len(Template_Level_Old)

		Dim The_Template_Level

		'�쿴�ڵ��Ƿ��ظ�
		The_Template_Level	= Template_Parent_ID_Arr(1) & Right(Template_Level, 2)
		Sql	= "Select [Template_ID], [Template_Level] From [Monolith_Template] Where [Template_ID]<> " & Template_ID & " And  [Template_Level] ='" & The_Template_Level & "'"
		Rs.Open Sql, Conn, 1, 1
		if Not (Rs.Bof Or Rs.Eof) then
			Response.write "��Ŀ����ȼ��ظ������������ã�"
			Response.end
		end if
		Rs.Close
			
		'�����¼��ڵ�
		Sql	= "Select [Template_ID], [Template_Level] From [Monolith_Template]"
		Rs.Open Sql, Conn, 1, 3
		Do While Not (Rs.Bof Or Rs.Eof) 
			if Left(Rs(1), Template_Level_Len) = Template_Level_Old then
				Sql	= "Update [Monolith_Template] Set "
				Sql = Sql & "[Template_Level]='"		& Template_Level & Right(Rs(1), Len(Rs(1)) - Template_Level_Len) & "' "
				Sql = Sql & " Where [Template_ID]=" & Rs(0)
				Conn.execute(Sql)
'				Response.write Sql & "<br>"
			end if
		  Rs.MoveNext
		Loop
		Rs.Close

		'������Ŀ
'		Sql	= "Update [Monolith_Template] Set "
'		Sql = Sql & "[Template_Parent_ID]='"	& Template_Parent_ID_Arr(0) 	& "', "
'		Sql = Sql & "[Template_Level]='" 			& Template_Level 			& "', "
'		Sql = Sql & "[Template_Title]='" 			& Template_Title 			& "', "
'		Sql = Sql & "[Template_HasChild]="		&	Template_HasChild		& " "
'		Sql = Sql & " Where [Template_ID]="		& Template_ID
'		Conn.execute(Sql)
		
		Dim strData
		Dim intFieldCount
		Dim i

		intFieldCount = Request.Form("hdnCount")
		For i=1 To intFieldCount
			strData = strData & Request.Form("hdnTemplate_Content" & i)
		Next
		
		Dim Template_Content
		Template_Content			= Replace(strData, "t extarea", "textarea")
		
		'-----------------------0-------------------1----------------2----------------3----------------------4
		Sql = "Select [Template_Parent_ID], [Template_Level], [Template_Title], [Template_HasChild], [Template_Content] From [Monolith_Template] Where [Template_ID]=" & Template_ID
		Rs.open Sql, conn , 1, 3
		if Not (Rs.Bof Or Rs.Eof) then
			Rs(0) = Template_Parent_ID_Arr(0)
			Rs(1) = Template_Level
			Rs(2) = Template_Title
			Rs(3)	= Template_HasChild
			Rs(4)	= Template_Content
			Rs.Update
		end if
		Rs.Close

		'�����丸��ĿΪ�нڵ��
		Sql	= "Update [Monolith_Template] Set "
		Sql = Sql & "[Template_HasChild]=true "
		Sql = Sql & " Where [Template_ID]=" & Template_Parent_ID_Arr(0)
		Conn.execute(Sql)

		'����ԭ����ĿΪ�нڵ��
		'����¼�û�нڵ��ˣ� ��ýڵ��ΪҶ��
		Sql	= "Select [Template_ID] From [Monolith_Template] Where [Template_Parent_ID]=" & Template_Parent_ID_Old
		Set	Rs	= Conn.execute(Sql)
		if Rs.Bof Or Rs.Eof then
			Sql	= "Update [Monolith_Template] Set "
			Sql = Sql & "[Template_HasChild]=false "
			Sql = Sql & " Where [Template_ID]=" & Template_Parent_ID_Old
			Conn.execute(Sql)
		end if
		Rs.Close

		Response.write "<script>alert('�޸ĳɹ���');</script>"

	end if
	

	'---------------------0----------------1-------------------2-----------------3-------------------4-------------------5
	Sql = "Select [Template_ID], [Template_Parent_ID], [Template_Level], [Template_Title], [Template_HasChild], [Template_Content] From [Monolith_Template] Where [Template_ID]= " & Template_ID
	Rs.Open Sql, Conn , 1, 1
	if Not (Rs.Bof Or Rs.Eof) then
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<form method="POST" action="<%=ThisPage%>?Edit=True&Template_ID=<%=Template_ID%>" onsubmit="fnPreHandle();return check();" name="FrmAddTemplate">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;ģ����� &gt;&gt; ģ���޸�</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">��ĿID��</font></td>
			<td bgcolor="#FFFFFF"><font color="#FF0000"><b><%=Rs(0)%></b></font></td>
		</tr>
		<tr <% 'if rs(4) then response.write "style="" display:none """  %>>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">����ĿID��</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="Template_Parent_ID" style="width=250px;" readonly class="InputBox">
			<option value="0|0">��Ϊ���ڵ�</option>
			<%
			'------------------0-------------1------------2
			Sql	= "Select [Template_ID], [Template_Level], [Template_Title] From [Monolith_Template] Order By [Template_Level]"
			RsTemp.Open Sql, Conn, 1, 1
			Do While Not (RsTemp.Bof Or RsTemp.Eof)
				if RsTemp(0) <> Rs(0) then
			%>
			<option value="<%=RsTemp(0) & "|" & RsTemp(1) %>" <% if rstemp(0) = rs(1) then response.write "selected" %>>
			<%LoopNBSP(Len(RsTemp(1)))%><%=RsTemp(0)%>-<%=RsTemp(2)%></option>
			<%
				end if
				RsTemp.MoveNext
			Loop
			RsTemp.Close
			%></select>
			<input type="hidden" name="Template_Parent_ID_Old" value="<%=Rs(1)%>">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr <%if rs(4) then response.write "style="" display:none """  %>>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">������������֣�����Բ����޸ģ�</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">��Ŀ����ȼ���</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="Template_Level" style="width=250px;" class="InputBox">
			<%
			Dim Same_Template_Level
			'------------------0
			Sql	= "Select [Template_Level] From [Monolith_Template] Order By [Template_Level]"
			RsTemp.Open Sql, Conn, 1, 1
			Do While Not (RsTemp.Bof Or RsTemp.Eof)
				if Len(RsTemp(0)) = Len(Rs(2)) then
					if Left(RsTemp(0), Len(RsTemp(0))-2) = Left(Rs(2), Len(Rs(2))-2) then
						if RsTemp(0) <> Rs(2) then
							Same_Template_Level	= Same_Template_Level & RsTemp(0) & ","
						end if
					end if
				end if
				RsTemp.MoveNext
			Loop
			RsTemp.Close
			i =1
			for i = 1 to 99
				if Len(i) = 1 then i = "0" & i
				The_Template_Level	= Left(Rs(2), Len(Rs(2))-2) & i
				if instr(Same_Template_Level, The_Template_Level) = 0 then
			%>
			<option value="<%=The_Template_Level%>" <% if the_template_level = rs(2) then response.write "selected" %>>
			<%=The_Template_Level%></option>
			<%
				end if
			Next
			%></select>
			<input type="hidden" name="Template_Level_Old" value="<%=Rs(2)%>">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">�� <%
			Sql	= "Select Count(*) As [Template_Level_Count] From [Monolith_Template] Where [Template_Level]='" & Rs(2) & "'"
			RsTemp.Open Sql, Conn, 1, 1
			if RsTemp(0) > 1 then
			%> <font color="#FF0000">��Ŀ����ȼ��ظ������޸�</font> <%
			else
				Response.write "��������"
			end if
			RsTemp.Close
			%> ��</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">��Ŀ���⣺</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Template_Title" size="40" style="width=250px;" value="<%=Rs(3)%>" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">������25�����ֻ�50��Ӣ����ĸ��Χ�ڣ�</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">�Ƿ��������Ŀ��</font></td>
			<td bgcolor="#FFFFFF">
			<input type="radio" value="true" name="Template_HasChild" id="Template_HasChild_1" <% if rs(4) = true then response.write "checked" %>><label for="Template_HasChild_1">�ǣ���Ϊ����Ŀ</label><input type="radio" value="false" <% if rs(4) = false then response.write "checked" %> name="Template_HasChild" id="Template_HasChild_0"><label for="Template_HasChild_0">����ΪҶ��</label></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">������Ϊ�ڵ���ѡ���ǣ�</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">ģ�����ݣ�</font></td>
			<td bgcolor="#FFFFFF">
			<textarea rows="8" name="Template_Content" id="Template_Content" cols="100" class="InputBox"><%=Replace(Rs(5), "textarea","t extarea")%></textarea></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">������Ϊ������Ŀ����Ҫ����ģ�����ݣ�</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">��</td>
			<td bgcolor="#FFFFFF">
			<input type="hidden" value="Same_Level" name="Template_Level_Code">
			<input type="submit" value=" �� �� (Alt + Y) " name="B1" class="InputBox" accesskey="Y">
			<input type="button" value=" Ԥ �� ģ �� (Alt + V) " name="B2" class="InputBox" onclick="runCode(Template_Content)" accesskey="V"><div id="divHidden">
			</div>
			</td>
		</tr>
	</form>
</table>
</form>
<%
	end if
	Set	RsTemp=nothing
	Call CloseRs
	Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->