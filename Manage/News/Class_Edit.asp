<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" --><!-- #include virtual ="/Manage/Check.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript">
function check()
{
  var Check_News_Class_Title	= document.FrmEditMenu.News_Class_Title;
  if (Check_News_Class_Title.value == "")
  {
 	alert("请输入栏目标题！");
	Check_News_Class_Title.focus();
	return false;
  }
}
</script>

<%
	Dim RsTemp
	Set	RsTemp				= Server.CreateObject("ADODB.RecordSet")

	Dim ThisPage
	ThisPage			   = Request.ServerVariables("PATH_INFO")

	Dim News_Class_ID, News_Class_Parent_ID, News_Class_Parent_ID_Arr,News_Class_Parent_ID_Old, News_Class_Level, News_Class_Level_Len, News_Class_Level_Old, News_Class_Title, News_Class_HasChild

	News_Class_ID						= Request("News_Class_ID")
	if Request("Edit") = "True" then
		News_Class_Parent_ID		= Request("News_Class_Parent_ID")
		News_Class_Level				= Request("News_Class_Level")
		News_Class_Level_Old		= Request("News_Class_Level_Old")
		News_Class_Title				= Request("News_Class_Title")
		News_Class_HasChild			= Request("News_Class_HasChild")
		News_Class_Parent_ID_Old=	Request("News_Class_Parent_ID_Old")
		
		News_Class_Parent_ID_Arr	= Split(News_Class_Parent_ID, "|")
		
		'跨段的时候用
		if News_Class_Parent_ID_Arr(1) <> 0 then
			if Len(News_Class_Level) + 2 <> Len(News_Class_Parent_ID_Arr(1)) then
				News_Class_Level	= News_Class_Parent_ID_Arr(1) & Right(News_Class_Level, 2)
			end if
		else
				News_Class_Parent_ID_Arr(1)	= ""
				News_Class_Level	= Right(News_Class_Level, 2)
		end if

		News_Class_Title				= MonolithEncode(News_Class_Title)
		News_Class_Level_Len	= Len(News_Class_Level_Old)

		Dim The_News_Class_Level

		'察看节点是否重复
		The_News_Class_Level	= News_Class_Parent_ID_Arr(1) & Right(News_Class_Level, 2)
		Sql	= "Select [News_Class_ID], [News_Class_Level] From [Monolith_News_Class] Where [News_Class_ID]<> " & News_Class_ID & " And  [News_Class_Level] ='" & The_News_Class_Level & "'"
		Rs.Open Sql, Conn, 1, 1
		if Not (Rs.Bof Or Rs.Eof) then
			Response.write "栏目排序等级重复，请重新设置！"
			Response.end
		end if
		Rs.Close
			
		'更新下级节点
		Sql	= "Select [News_Class_ID], [News_Class_Level] From [Monolith_News_Class]"
		Rs.Open Sql, Conn, 1, 3
		Do While Not (Rs.Bof Or Rs.Eof) 
			if Left(Rs(1), News_Class_Level_Len) = News_Class_Level_Old then
				Sql	= "Update [Monolith_News_Class] Set "
				Sql = Sql & "[News_Class_Level]='"		& News_Class_Level & Right(Rs(1), Len(Rs(1)) - News_Class_Level_Len) & "' "
				Sql = Sql & " Where [News_Class_ID]=" & Rs(0)
				Conn.execute(Sql)
'				Response.write Sql & "<br>"
			end if
		  Rs.MoveNext
		Loop
		Rs.Close

		'更新栏目
		Sql	= "Update [Monolith_News_Class] Set "
		Sql = Sql & "[News_Class_Parent_ID]='"	& News_Class_Parent_ID_Arr(0) 	& "', "
		Sql = Sql & "[News_Class_Level]='" 			& News_Class_Level 			& "', "
		Sql = Sql & "[News_Class_Title]='" 			& News_Class_Title 			& "', "
		Sql = Sql & "[News_Class_HasChild]="		&	News_Class_HasChild		& " "
		Sql = Sql & " Where [News_Class_ID]="		& News_Class_ID
		Conn.execute(Sql)

		'更新其父栏目为有节点的
		Sql	= "Update [Monolith_News_Class] Set "
		Sql = Sql & "[News_Class_HasChild]=true "
		Sql = Sql & " Where [News_Class_ID]=" & News_Class_Parent_ID_Arr(0)
		Conn.execute(Sql)

		'更新原父栏目为有节点的
		'如果下级没有节点了， 则该节点变为叶子
		Sql	= "Select [News_Class_ID] From [Monolith_News_Class] Where [News_Class_Parent_ID]=" & News_Class_Parent_ID_Old
		Set	Rs	= Conn.execute(Sql)
		if Rs.Bof Or Rs.Eof then
			Sql	= "Update [Monolith_News_Class] Set "
			Sql = Sql & "[News_Class_HasChild]=false "
			Sql = Sql & " Where [News_Class_ID]=" & News_Class_Parent_ID_Old
			Conn.execute(Sql)
		end if
		Rs.Close

		Response.write "<font color=""#FF0000"">修改成功！</font>"
	end if
	

	'---------------------0-------------------1--------------------2-------------------3----------------------4
	Sql = "Select [News_Class_ID], [News_Class_Parent_ID], [News_Class_Level], [News_Class_Title], [News_Class_HasChild] From [Monolith_News_Class] Where [News_Class_ID]= " & News_Class_ID
	Rs.Open Sql, Conn , 1, 1
	if Not (Rs.Bof Or Rs.Eof) then
%>
<form method="POST" action="<%=ThisPage%>?Edit=True&News_Class_ID=<%=News_Class_ID%>"  onSubmit="return check()" name="FrmEditMenu">
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;新闻栏目管理 &gt;&gt; 栏目修改</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">栏目ID：</font></td>
			<td bgcolor="#F6F6F6"><font color="#FF0000"><b><%=Rs(0)%></b></font></td>
		</tr>
		<tr <% 'if Rs(4) then Response.write "style="" display:none """  %>>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">父栏目ID：</font></td>
			<td bgcolor="#F6F6F6">
			<select size="1" name="News_Class_Parent_ID" style="width=250px;" readonly>
			<option value="0|0">作为根节点</option>
			<%
			'------------------0-------------1------------2
			Sql	= "Select [News_Class_ID], [News_Class_Level], [News_Class_Title] From [Monolith_News_Class] Order By [News_Class_Level]"
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
			<input type="hidden" name="News_Class_Parent_ID_Old" value="<%=Rs(1)%>">
			<font color="#FF0000">*</font></td>
		</tr>
		<tr <%if Rs(4) then Response.write "style="" display:none """  %>>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#F6F6F6">（如果存在数字，则可以不用修改）</td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">栏目排序等级：</font></td>
			<td bgcolor="#F6F6F6"><select size="1" name="News_Class_Level" style="width=250px;">
			<%
			Dim Same_News_Class_Level, i
			'------------------0
			Sql	= "Select [News_Class_Level] From [Monolith_News_Class] Order By [News_Class_Level]"
			RsTemp.Open Sql, Conn, 1, 1
			Do While Not (RsTemp.Bof Or RsTemp.Eof)
				if Len(RsTemp(0)) = Len(Rs(2)) then
					if Left(RsTemp(0), Len(RsTemp(0))-2) = Left(Rs(2), Len(Rs(2))-2) then
						if RsTemp(0) <> Rs(2) then
							Same_News_Class_Level	= Same_News_Class_Level & RsTemp(0) & ","
						end if
					end if
				end if
				RsTemp.MoveNext
			Loop
			RsTemp.Close
			i =1
			for i = 1 to 99
				if Len(i) = 1 then i = "0" & i
				The_News_Class_Level	= Left(Rs(2), Len(Rs(2))-2) & i
				if instr(Same_News_Class_Level, The_News_Class_Level) = 0 then
			%>
			<option value="<%=The_News_Class_Level%>" <% if The_News_Class_Level = Rs(2) then Response.write "selected" %>><%=The_News_Class_Level%></option>
			<%
				end if
			Next
			%>
				</select>
				<input type="hidden" name="News_Class_Level_Old" value="<%=Rs(2)%>">
				<font color="#FF0000">*</font></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#F6F6F6">（
			<%
			Sql	= "Select Count(*) As [News_Class_Level_Count] From [Monolith_News_Class] Where [News_Class_Level]='" & Rs(2) & "'"
			RsTemp.Open Sql, Conn, 1, 1
			if RsTemp(0) > 1 then
			%>
			<font color="#FF0000">栏目排序等级重复，请修改</font>
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
			<font color="#FFFFFF">栏目标题：</font></td>
			<td bgcolor="#F6F6F6">
			<input type="text" name="News_Class_Title" size="40" style="width=250px;" value="<%=Rs(3)%>">
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
			<input type="radio" value="true" name="News_Class_HasChild" id="News_Class_HasChild_1" <% if Rs(4) = true then Response.write "checked" %>><label for="News_Class_HasChild_1">是，做为子栏目</label><input type="radio" value="false" <% if Rs(4) = false then Response.write "checked" %>  name="News_Class_HasChild" id="News_Class_HasChild_0"><label for="News_Class_HasChild_0">否，作为叶子</label></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#F6F6F6">（如作为节点请选择是）</td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#F6F6F6"><input type="hidden" value="Same_Level" name="News_Class_Level_Code"><input type="submit" value=" 提 交 " name="B1">
			<input type="reset" value=" 重 置 " name="B2"></td>
		</tr>
	</table>
</form>
<%
	end if
	Set	RsTemp=nothing
	Call CloseRs
	Call CloseConn
%><!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->