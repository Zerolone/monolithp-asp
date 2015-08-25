<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript">
function check()
{
  var Check_Tree_Title	= document.FrmEditMenu.Tree_Title;
  if (Check_Tree_Title.value == "")
  {
 	alert("请输入菜单标题！");
	Check_Tree_Title.focus();
	return false;
  }
}
</script>

<%
	Dim RsTemp
	Set	RsTemp				= Server.CreateObject("ADODB.RecordSet")


	Dim Tree_ID, Tree_Parent_ID, Tree_Parent_ID_Arr,Tree_Parent_ID_Old, Tree_Level, Tree_Level_Len, Tree_Level_Old, Tree_Title, Tree_Url, Tree_Target, Tree_HasChild

	Tree_ID						= Request("Tree_ID")
	if Request("Edit") = "True" then
		Tree_Parent_ID		= Request("Tree_Parent_ID")
		Tree_Level				= Request("Tree_Level")
		Tree_Level_Old		= Request("Tree_Level_Old")
		Tree_Title				= Request("Tree_Title")
		Tree_Url					= Request("Tree_Url")
		Tree_Target				= Request("Tree_Target")
		Tree_HasChild			= Request("Tree_HasChild")
		Tree_Parent_ID_Old=	Request("Tree_Parent_ID_Old")
		
		Tree_Parent_ID_Arr	= Split(Tree_Parent_ID, "|")
		
		'跨段的时候用
		if Tree_Parent_ID_Arr(1) <> 0 then
			if Len(Tree_Level) + 2 <> Len(Tree_Parent_ID_Arr(1)) then
				Tree_Level	= Tree_Parent_ID_Arr(1) & Right(Tree_Level, 2)
			end if
		else
				Tree_Parent_ID_Arr(1)	= ""
				Tree_Level	= Right(Tree_Level, 2)
		end if

		Tree_Title				= MonolithEncode(Tree_Title)
		Tree_Url					= MonolithEncode(Tree_Url)

		Tree_Level_Len	= Len(Tree_Level_Old)

		Dim The_Tree_Level

		'察看节点是否重复
		The_Tree_Level	= Tree_Parent_ID_Arr(1) & Right(Tree_Level, 2)
		Sql	= "Select [Tree_ID], [Tree_Level] From [Monolith_Tree] Where [Tree_ID]<> " & Tree_ID & " And  [Tree_Level] ='" & The_Tree_Level & "'"
		Rs.Open Sql, Conn, 1, 1
		if Not (Rs.Bof Or Rs.Eof) then
			Response.write "菜单排序等级重复，请重新设置！"
			Response.end
		end if
		Rs.Close
			
		'更新下级节点
		Sql	= "Select [Tree_ID], [Tree_Level] From [Monolith_Tree]"
		Rs.Open Sql, Conn, 1, 3
		Do While Not (Rs.Bof Or Rs.Eof) 
			if Left(Rs(1), Tree_Level_Len) = Tree_Level_Old then
				Sql	= "Update [Monolith_Tree] Set "
				Sql = Sql & "[Tree_Level]='"		& Tree_Level & Right(Rs(1), Len(Rs(1)) - Tree_Level_Len) & "' "
				Sql = Sql & " Where [Tree_ID]=" & Rs(0)
				Conn.execute(Sql)
'				Response.write Sql & "<br>"
			end if
		  Rs.MoveNext
		Loop
		Rs.Close

		'更新菜单
		Sql	= "Update [Monolith_Tree] Set "
		Sql = Sql & "[Tree_Parent_ID]='"	& Tree_Parent_ID_Arr(0) 	& "', "
		Sql = Sql & "[Tree_Level]='" 			& Tree_Level 			& "', "
		Sql = Sql & "[Tree_Title]='" 			& Tree_Title 			& "', "
		Sql = Sql & "[Tree_Url]='" 				& Tree_Url 				& "', "
		Sql = Sql & "[Tree_Target]='" 		& Tree_Target			& "', "
		Sql = Sql & "[Tree_HasChild]="		&	Tree_HasChild		& " "
		Sql = Sql & " Where [Tree_ID]="		& Tree_ID
		Conn.execute(Sql)

		'更新其父菜单为有节点的
		Sql	= "Update [Monolith_Tree] Set "
		Sql = Sql & "[Tree_HasChild]=true "
		Sql = Sql & " Where [Tree_ID]=" & Tree_Parent_ID_Arr(0)
		Conn.execute(Sql)

		'更新原父菜单为有节点的
		'如果下级没有节点了， 则该节点变为叶子
		Sql	= "Select [Tree_ID] From [Monolith_Tree] Where [Tree_Parent_ID]=" & Tree_Parent_ID_Old
		Set	Rs	= Conn.execute(Sql)
		if Rs.Bof Or Rs.Eof then
			Sql	= "Update [Monolith_Tree] Set "
			Sql = Sql & "[Tree_HasChild]=false "
			Sql = Sql & " Where [Tree_ID]=" & Tree_Parent_ID_Old
			Conn.execute(Sql)
		end if
		Rs.Close
		
		Response.write("<script language='javascript'>alert('修改成功！')</script>")
		
	end if

	'-----------------0-------------1----------------2------------3--------------4----------5---------------6-----------------
	Sql = "Select [Tree_ID], [Tree_Parent_ID], [Tree_Level], [Tree_Title], [Tree_Url], [Tree_Target], [Tree_HasChild] From [Monolith_Tree] Where [Tree_ID]= " & Tree_ID
	Rs.Open Sql, Conn , 1, 1
	if Not (Rs.Bof Or Rs.Eof) then
%>
<form method="POST" action="Edit.asp?Edit=True&Tree_ID=<%=Tree_ID%>"  onSubmit="return check()" name="FrmEditMenu">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;左侧菜单管理 &gt;&gt; 菜单修改</font></b></td>
		</tr>		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">菜&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 单&nbsp; ID：</font></td>
			<td bgcolor="#FFFFFF"><font color="#FF0000"><b><%=Rs(0)%></b></font></td>
		</tr>
		<tr <% 'if Rs(6) then Response.write "style="" display:none """  %>>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">父&nbsp; 菜&nbsp; 单&nbsp; ID：</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="Tree_Parent_ID" style="width=250px;" readonly class="InputBox">
			<option value="0|0">作为根节点</option>
			<%
			'------------------0-------------1------------2
			Sql	= "Select [Tree_ID], [Tree_Level], [Tree_Title] From [Monolith_Tree] Order By [Tree_Level]"
			RsTemp.Open Sql, Conn, 1, 1
			Do While Not (RsTemp.Bof Or RsTemp.Eof)
				if RsTemp(0) <> Rs(0) then
			%>
			<option value="<%=RsTemp(0) & "|" & RsTemp(1) %>" <% if RsTemp(0) = Rs(1) then Response.write "selected" %>><%LoopNBSP(Len(RsTemp(1)))%><%=RsTemp(0)%>-<%=RsTemp(2)%></option>
			<%
				end if
				RsTemp.MoveNext
			Loop
			RsTemp.Close
			%>
			</select>
			<input type="hidden" name="Tree_Parent_ID_Old" value="<%=Rs(1)%>">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr <%if Rs(6) then Response.write "style="" display:none """  %>>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（如果存在数字，则可以不用修改）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">菜 单 排序等级：</font></td>
			<td bgcolor="#FFFFFF"><select size="1" name="Tree_Level" style="width=250px;" class="InputBox">
			<%
			Dim Same_Tree_Level, i
			'------------------0
			Sql	= "Select [Tree_Level] From [Monolith_Tree] Order By [Tree_Level]"
			RsTemp.Open Sql, Conn, 1, 1
			Do While Not (RsTemp.Bof Or RsTemp.Eof)
				if Len(RsTemp(0)) = Len(Rs(2)) then
					if Left(RsTemp(0), Len(RsTemp(0))-2) = Left(Rs(2), Len(Rs(2))-2) then
						if RsTemp(0) <> Rs(2) then
							Same_Tree_Level	= Same_Tree_Level & RsTemp(0) & ","
						end if
					end if
				end if
				RsTemp.MoveNext
			Loop
			RsTemp.Close
			i =1
			for i = 1 to 99
				if Len(i) = 1 then i = "0" & i
				The_Tree_Level	= Left(Rs(2), Len(Rs(2))-2) & i
				if instr(Same_Tree_Level, The_Tree_Level) = 0 then
			%>
			<option value="<%=The_Tree_Level%>" <% if The_Tree_Level = Rs(2) then Response.write "selected" %>><%=The_Tree_Level%></option>
			<%
				end if
			Next
			%>
				</select>
				<input type="hidden" name="Tree_Level_Old" value="<%=Rs(2)%>">
				<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（
			<%
			Sql	= "Select Count(*) As [Tree_Level_Count] From [Monolith_Tree] Where [Tree_Level]='" & Rs(2) & "'"
			RsTemp.Open Sql, Conn, 1, 1
			if RsTemp(0) > 1 then
			%>
			<font color="#FF0000">菜单排序等级重复，请修改</font>
			<%
			else
				Response.write "排序正常"
			end if
			RsTemp.Close
			%>
			）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">菜&nbsp; 单&nbsp; 标&nbsp; 题：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Tree_Title" size="40" style="width=250px;" value="<%=Rs(3)%>" class="InputBox">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（请在25个汉字或50个英文字母范围内）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">菜&nbsp; 单&nbsp; 链&nbsp; 接：</font></td>
			<td bgcolor="#FFFFFF">
			<input type="text" name="Tree_Url" size="40" style="width=250px;" value="<%=Rs(4)%>" class="InputBox"></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（如作为节点可以不填写）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">菜 单 链接框架：</font></td>
			<td bgcolor="#FFFFFF"><select size="1" name="Tree_Target" style="width=250px;" class="InputBox">
			<option value="tabWin" 	<% if Rs(5) = "tabWin" 	then Response.write "selected" %> >默认窗口</option>
			<option value="_self"		<% if Rs(5) = "_self"	 	then Response.write "selected" %> >相同框架</option>
			<option value="_top"		<% if Rs(5) = "_top" 		then Response.write "selected" %> >整页</option>
			<option value="_blank"	<% if Rs(5) = "_blank" 	then Response.write "selected" %> >新窗口</option>
			<option value="_parent"	<% if Rs(5) = "_parent" then Response.write "selected" %> >父窗口</option>
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
			<input type="radio" value="true" name="Tree_HasChild" id="Tree_HasChild_1" <% if Rs(6) = true then Response.write "checked" %>><label for="Tree_HasChild_1">是，做为子菜单</label><input type="radio" value="false" <% if Rs(6) = false then Response.write "checked" %>  name="Tree_HasChild" id="Tree_HasChild_0"><label for="Tree_HasChild_0">否，作为叶子</label></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">（如作为节点请选择是）</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF"><input type="hidden" value="Same_Level" name="Tree_Level_Code">
			<input type="submit" value=" 修 改 (Alt + Y) " name="B1" class="InputBox" accesskey="Y">
			<input type="button" value=" 取 消 (Alt + N) " name="B2" onclick="window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);" class="InputBox" accesskey="N"></td>
		</tr>
	</table>
</form>
<%
	end if
	Set	RsTemp=nothing
	Call CloseRs
	Call CloseConn
%>
<table><tr><td height="160"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->