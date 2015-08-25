<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript">
function check()
{
  var Check_Info_Class_Title	= document.FrmEditMenu.Info_Class_Title;
  if (Check_Info_Class_Title.value == "")
  {
 	alert("请输入栏目标题！");
	Check_Info_Class_Title.focus();
	return false;
  }
}

function CheckBoxSelect(CheckBoxName, Code)
{
	switch (Code){
	case 1:
	  for(i=0; i<= CheckBoxName.length-1;i++)
	  {
	  	CheckBoxName(i).checked=true;
	  }
	  break;

	case 2:
	  for(i=0; i<= CheckBoxName.length-1;i++)
	  {
	  	CheckBoxName(i).checked=CheckBoxName(i).checked==true?false:true;
	  }
	  break;

	case 3:
	  for(i=0; i<= CheckBoxName.length-1;i++)
	  {
	  	CheckBoxName(i).checked=false;
	  }
	  break;

	 }
}
</script>

<%
	Dim RsTemp
	Set	RsTemp				= Server.CreateObject("ADODB.RecordSet")

	Dim ThisPage
	ThisPage			   = Request.ServerVariables("PATH_INFO")

	Dim Info_Class_ID, Info_Class_Parent_ID, Info_Class_Parent_ID_Arr,Info_Class_Parent_ID_Old, Info_Class_Level, Info_Class_Level_Len, Info_Class_Level_Old, Info_Class_Title, Info_Class_HasChild
	Dim Info_Users_Level

	Info_Class_ID						= Request("Info_Class_ID")
	if Request("Edit") = "True" then
		Info_Class_Parent_ID		= Request("Info_Class_Parent_ID")
		Info_Class_Level				= Request("Info_Class_Level")
		Info_Class_Level_Old		= Request("Info_Class_Level_Old")
		Info_Class_Title				= Request("Info_Class_Title")
		Info_Class_HasChild			= Request("Info_Class_HasChild")
		Info_Class_Parent_ID_Old=	Request("Info_Class_Parent_ID_Old")
		Info_Users_Level				= Request("Info_Users_Level")
		
		Info_Class_Parent_ID_Arr	= Split(Info_Class_Parent_ID, "|")
		
		'跨段的时候用
		if Info_Class_Parent_ID_Arr(1) <> 0 then
			if Len(Info_Class_Level) + 2 <> Len(Info_Class_Parent_ID_Arr(1)) then
				Info_Class_Level	= Info_Class_Parent_ID_Arr(1) & Right(Info_Class_Level, 2)
			end if
		else
				Info_Class_Parent_ID_Arr(1)	= ""
				Info_Class_Level	= Right(Info_Class_Level, 2)
		end if

		Info_Class_Title				= MonolithEncode(Info_Class_Title)
		Info_Class_Level_Len	= Len(Info_Class_Level_Old)

		Dim The_Info_Class_Level

		'察看节点是否重复
		The_Info_Class_Level	= Info_Class_Parent_ID_Arr(1) & Right(Info_Class_Level, 2)
		Sql	= "Select [Info_Class_ID], [Info_Class_Level] From [Monolith_Info_Class] Where [Info_Class_ID]<> " & Info_Class_ID & " And  [Info_Class_Level] ='" & The_Info_Class_Level & "'"
		Rs.Open Sql, Conn, 1, 1
		if Not (Rs.Bof Or Rs.Eof) then
			Response.write "栏目排序等级重复，请重新设置！"
			Response.end
		end if
		Rs.Close
			
		'更新下级节点
		Sql	= "Select [Info_Class_ID], [Info_Class_Level] From [Monolith_Info_Class]"
		Rs.Open Sql, Conn, 1, 3
		Do While Not (Rs.Bof Or Rs.Eof) 
			if Left(Rs(1), Info_Class_Level_Len) = Info_Class_Level_Old then
				Sql	= "Update [Monolith_Info_Class] Set "
				Sql = Sql & "[Info_Class_Level]='"		& Info_Class_Level & Right(Rs(1), Len(Rs(1)) - Info_Class_Level_Len) & "' "
				Sql = Sql & " Where [Info_Class_ID]=" & Rs(0)
				Conn.execute(Sql)
'				Response.write Sql & "<br>"
			end if
		  Rs.MoveNext
		Loop
		Rs.Close

		'更新栏目
		Sql	= "Update [Monolith_Info_Class] Set "
		Sql = Sql & "[Info_Class_Parent_ID]='"	& Info_Class_Parent_ID_Arr(0) 	& "', "
		Sql = Sql & "[Info_Class_Level]='" 			& Info_Class_Level 			& "', "
		Sql = Sql & "[Info_Class_Title]='" 			& Info_Class_Title 			& "', "
		Sql = Sql & "[Info_Users_Level]='" 			& Info_Users_Level 			& "', "

		Sql = Sql & "[Info_Class_HasChild]="		&	Info_Class_HasChild		& " "
		Sql = Sql & " Where [Info_Class_ID]="		& Info_Class_ID
		Conn.execute(Sql)

		'更新其父栏目为有节点的
		Sql	= "Update [Monolith_Info_Class] Set "
		Sql = Sql & "[Info_Class_HasChild]=true "
		Sql = Sql & " Where [Info_Class_ID]=" & Info_Class_Parent_ID_Arr(0)
		Conn.execute(Sql)

		'更新原父栏目为有节点的
		'如果下级没有节点了， 则该节点变为叶子
		Sql	= "Select [Info_Class_ID] From [Monolith_Info_Class] Where [Info_Class_Parent_ID]=" & Info_Class_Parent_ID_Old
		Set	Rs	= Conn.execute(Sql)
		if Rs.Bof Or Rs.Eof then
			Sql	= "Update [Monolith_Info_Class] Set "
			Sql = Sql & "[Info_Class_HasChild]=false "
			Sql = Sql & " Where [Info_Class_ID]=" & Info_Class_Parent_ID_Old
			Conn.execute(Sql)
		end if
		Rs.Close

		Response.write("<script language='javascript'>alert('修改成功！')</script>")
	end if
	

	'---------------------0-------------------1--------------------2-------------------3----------------------4----------------------5
	Sql = "Select [Info_Class_ID], [Info_Class_Parent_ID], [Info_Class_Level], [Info_Class_Title], [Info_Class_HasChild], [Info_Users_Level] From [Monolith_Info_Class] Where [Info_Class_ID]= " & Info_Class_ID
	Rs.Open Sql, Conn , 1, 1
	if Not (Rs.Bof Or Rs.Eof) then
%>
<form method="POST" action="<%=ThisPage%>?Edit=True&Info_Class_ID=<%=Info_Class_ID%>" onsubmit="return check()" name="FrmEditMenu">
	<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;信息栏目管理 &gt;&gt; 栏目修改</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">栏目ID：</font></td>
			<td bgcolor="#F6F6F6"><font color="#FF0000"><b><%=Rs(0)%></b></font></td>
		</tr>
		<tr <% 'if rs(4) then response.write "style="" display:none """  %>>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">父栏目ID：</font></td>
			<td bgcolor="#F6F6F6">
			<select size="1" name="Info_Class_Parent_ID" style="width=250px;" class="InputBox">
			<option value="0|0">作为根节点</option>
			<%
			'------------------------0----------------1-------------------2
			Sql	= "Select [Info_Class_ID], [Info_Class_Level], [Info_Class_Title]  From [Monolith_Info_Class] Order By [Info_Class_Level]"
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
			<input type="hidden" name="Info_Class_Parent_ID_Old" value="<%=Rs(1)%>">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr <%if rs(4) then response.write "style="" display:none """  %>>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#F6F6F6">（如果存在数字，则可以不用修改）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">栏目排序等级：</font></td>
			<td bgcolor="#F6F6F6">
			<select size="1" name="Info_Class_Level" style="width=250px;" class="InputBox">
			<%
			Dim Same_Info_Class_Level, i
			'------------------0
			Sql	= "Select [Info_Class_Level] From [Monolith_Info_Class] Order By [Info_Class_Level]"
			RsTemp.Open Sql, Conn, 1, 1
			Do While Not (RsTemp.Bof Or RsTemp.Eof)
				if Len(RsTemp(0)) = Len(Rs(2)) then
					if Left(RsTemp(0), Len(RsTemp(0))-2) = Left(Rs(2), Len(Rs(2))-2) then
						if RsTemp(0) <> Rs(2) then
							Same_Info_Class_Level	= Same_Info_Class_Level & RsTemp(0) & ","
						end if
					end if
				end if
				RsTemp.MoveNext
			Loop
			RsTemp.Close
			i =1
			for i = 1 to 99
				if Len(i) = 1 then i = "0" & i
				The_Info_Class_Level	= Left(Rs(2), Len(Rs(2))-2) & i
				if instr(Same_Info_Class_Level, The_Info_Class_Level) = 0 then
			%>
			<option value="<%=The_Info_Class_Level%>" <% if the_info_class_level = rs(2) then response.write "selected" %>>
			<%=The_Info_Class_Level%></option>
			<%
				end if
			Next
			%></select>
			<input type="hidden" name="Info_Class_Level_Old" value="<%=Rs(2)%>">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#F6F6F6">（ <%
			Sql	= "Select Count(*) As [Info_Class_Level_Count] From [Monolith_Info_Class] Where [Info_Class_Level]='" & Rs(2) & "'"
			RsTemp.Open Sql, Conn, 1, 1
			if RsTemp(0) > 1 then
			%> <font color="#FF0000">栏目排序等级重复，请修改</font> <%
			else
				Response.write "排序正常"
			end if
			RsTemp.Close
			%> ）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">栏目标题：</font></td>
			<td bgcolor="#F6F6F6">
			<input type="text" name="Info_Class_Title" size="40" style="width=250px;" value="<%=Rs(3)%>" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#F6F6F6">（请在25个汉字或50个英文字母范围内）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">是否存在子栏目：</font></td>
			<td bgcolor="#F6F6F6">
			<input type="radio" value="true" name="Info_Class_HasChild" id="Info_Class_HasChild_1" <% if rs(4) = true then response.write "checked" %>><label for="Info_Class_HasChild_1">是，做为子栏目</label><input type="radio" value="false" <% if rs(4) = false then response.write "checked" %> name="Info_Class_HasChild" id="Info_Class_HasChild_0"><label for="Info_Class_HasChild_0">否，作为叶子</label></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#F6F6F6">（如作为节点请选择是）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">栏目浏览权限：</font></td>
			<td bgcolor="#F6F6F6"><%
			Dim Rs1
			Set Rs1		= Server.CreateObject("ADODB.Recordset")
			'-------------------0-----------1
			Sql = "Select [Level_ID], [Level_Name] From [Monolith_Users_Level] Order By [Level_Order] Asc"
			Rs1.Open Sql, Conn ,1, 1
			Do While Not (Rs1.Bof Or Rs1.Eof)
		%> <input type="checkbox" name="Info_Users_Level" value="<%=Rs1(0)%>" id="Info_Users_Level_<%=Rs1(0)%>" 
		<%
			Dim Info_Users_Level_Arr
			Info_Users_Level_Arr	= Split(Rs(5), ",")
			For i = LBound(Info_Users_Level_Arr) to UBound(Info_Users_Level_Arr)
//			  Response.write Info_Users_Level_Arr(i)
				if Cint(Trim(Info_Users_Level_Arr(i))) = Rs1(0) then Response.write "checked"
			Next
		%>
		><label for="Info_Users_Level_<%=Rs1(0)%>"><%=Rs1(1)%></label>
			<%
				Rs1.MoveNext
			Loop
			Rs1.Close
			Set Rs1=Nothing
		%>　
		<input type="checkbox" name="Info_Users_Level" value="0" id="Info_Users_Level_0" ><label for="Info_Users_Level_0">非会员</label>

		
		</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#F6F6F6">
			<input type="button" value=" 全 选 (Alt + A) " onclick="CheckBoxSelect(Info_Users_Level, 1)" class="InputBox" accesskey="A">
			<input type="button" value=" 反 选 (Alt + I) " onclick="CheckBoxSelect(Info_Users_Level, 2)" class="InputBox" accesskey="I">
			<input type="button" value=" 全 不 选 (Alt + N) " onclick="CheckBoxSelect(Info_Users_Level, 3)" class="InputBox" accesskey="N">
			</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#F6F6F6">
			<input type="hidden" value="Same_Level" name="Info_Class_Level_Code">
			<input type="submit" name="Submit0" value=" 修 改 (Alt + Y) " class="InputBox" accesskey="Y">
			<input type="button" value=" 取 消 (Alt + N) " name="B2" onclick="window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);" class="InputBox">
			</td>
		</tr>
	</table>
</form>
<%
	end if
	Set	RsTemp=nothing
	Call CloseRs
	Call CloseConn
%><table>
	<tr>
		<td height="200"></td>
	</tr>
</table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->