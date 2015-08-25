<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" --><!-- #include virtual ="/Manage/Check.asp"--><%
	Dim View_01, View_02, View_03, View_04, View_05, View_06, View_07, TheString
	
	View_01	= Request("View_01")
	View_02	= Request("View_02")
	View_03	= Request("View_03")
	View_04	= Request("View_04")
	View_05	= Request("View_05")
	View_06	= Request("View_06")
	View_07	= Request("View_07")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
<title>无标题文档</title>
</head>

<script language="javascript" src="/Js/All.js"></script>
<body>

<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<form method="POST" action="View.asp" name="FrmView">
		<tr bgcolor="#6A6A6A" height="26">
			<td colspan="2"><b><font color="#FFFFFF">&nbsp;公告管理 &gt;&gt; 公告预览</font></b></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">公告显示方式：</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="View_01" style="width=200" class="InputBox">
			<option value="1" <%if view_01=1 then response.write "selected"%>>向左滚动
			</option>
			<option value="2" <%if view_01=2 then response.write "selected"%>>向右滚动
			</option>
			<option value="3" <%if view_01=3 then response.write "selected"%>>向上滚动
			</option>
			<option value="4" <%if view_01=4 then response.write "selected"%>>向下滚动
			</option>
			</select></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">公告滚动条数：</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="View_02" style="width=200" class="InputBox"><%
				Dim i
				For i=1 to 20
			%>
			<option value="<%=i%>" <%if cint(view_02)=i then response.write "selected"%>>
			<%=i%></option>
			<%Next%></select></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">滚动位置宽度：</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="View_03" style="width=200" class="InputBox"><%For i=5 to 1000 step 5 %>
			<option value="<%=i%>" <%if cint(view_03)=i then response.write "selected"%>>
			<%=i%></option>
			<%Next%></select></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">滚动位置长度：</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="View_04" style="width=200" class="InputBox"><%For i=5 to 800 step 5 %>
			<option value="<%=i%>" <%if cint(view_04)=i then response.write "selected"%>>
			<%=i%></option>
			<%Next%></select></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">滚动速度：</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="View_05" style="width=200" class="InputBox"><%For i=1 to 10 %>
			<option value="<%=i%>" <%if cint(view_05)=i then response.write "selected"%>>
			<%=i%></option>
			<%Next%></select></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">是否显示日期：</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="View_06" style="width=200" class="InputBox">
			<option value="1" <%if cint(view_06)=1 then response.write "selected"%>>显示</option>
			<option value="2" <%if cint(view_06)=2 then response.write "selected"%>>不显示</option>
			</select></td>
		</tr>
		<tr>
			<td align="right" width="100" bgcolor="#999999" height="20">
			<font color="#FFFFFF">背景颜色：</font></td>
			<td bgcolor="#FFFFFF">
			<select size="1" name="View_07" style="width=200" class="InputBox">
			<option style="background-color:#FFFFFF;color: #FFFFFF" value="#FFFFFF">
			#FFFFFF</option>
			<option style="background-color:#008000;color: #008000" value="#008000">
			#008000</option>
			<option style="background-color:#800000;color: #800000" value="#800000">
			#800000</option>
			<option style="background-color:#808000;color: #808000" value="#808000">
			#808000</option>
			<option style="background-color:#000080;color: #000080" value="#000080">
			#000080</option>
			<option style="background-color:#800080;color: #800080" value="#800080">
			#800080</option>
			<option style="background-color:#808080;color: #808080" value="#808080">
			#808080</option>
			<option style="background-color:#FFFF00;color: #FFFF00" value="#FFFF00">
			#FFFF00</option>
			<option style="background-color:#00FF00;color: #00FF00" value="#00FF00">
			#00FF00</option>
			<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">
			#00FFFF</option>
			<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">
			#FF00FF</option>
			<option style="background-color:#C0C0C0;color: #C0C0C0" value="#C0C0C0">
			#C0C0C0</option>
			<option style="background-color:#FF0000;color: #FF0000" value="#FF0000">
			#FF0000</option>
			<option style="background-color:#0000FF;color: #0000FF" value="#0000FF">
			#0000FF</option>
			<option style="background-color:#008080;color: #008080" value="#008080">
			#008080</option>
			</select></td>
		</tr>
		<tr>
			<td width="100" bgcolor="#999999" height="20">　</td>
			<td bgcolor="#FFFFFF">
			<input type="submit" value=" 提 交 (Alt + S) " name="B1" class="InputBox" accesskey="A"></td>
		</tr>
	</form>
</table>
<%
	if View_01 <> "" then
	

		'---------------------------------------0---------------1-----------------2
		Sql = "Select Top " & View_02 & " [Bulletin_ID], [Bulletin_Title], [Bulletin_Date] From [Monolith_Bulletin] Order By [Bulletin_ID] Desc"
		Rs.Open Sql, Conn, 1, 1
		Do While Not(Rs.Bof Or Rs.Eof)
      TheString = TheString + "<A href=/Bulletin/Bulletin.asp?Bulletin_ID=" & Rs(0) & "  target='_blank'>" & rs(1)
      
      if View_06 = 1 then TheString = TheString & "&nbsp;[" & Rs(2) & "]"
      
      TheString = TheString & "</A>"
      
      if View_01 > 2 then
        TheString = TheString & "<BR><BR>"
      else
      	TheString = TheString & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
      end if
			Rs.MoveNext
		Loop
		Call CloseRs
		Call CloseConn

   Select Case View_01
      Case "1"
        View_01 = "left"
      Case "2"  
        View_01 = "right"
      Case "3"
        View_01 = "up"
      Case "4"
        View_01 = "down"
   End Select	

%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
	<tr bgcolor="#6A6A6A" height="26">
		<td><b><font color="#FFFFFF">&nbsp;公告管理 &gt;&gt; 公告预览区</font></b></td>
	</tr>
	<tr height="20">
		<td width="100" bgcolor="#FFFFFF">
		<marquee onmouseover="this.scrollAmount=0" onmouseout="this.scrollAmount=<%=View_05%>" scrollamount="<%=View_05%>" direction="<%=View_01%>" height="<%=View_03%>" width="<%=View_04%>" bgcolor="<%=View_07%>"><%=TheString%></marquee></td>
		</td>
	</tr>
</table>
<%
	View_03	= View_03 + 32
	end if
%>
<table><tr><td height="<%=302-View_03%>"></td></tr></table>
<!-- #include virtual ="/Inc/Monolith_ThisPage.asp" -->
</body>

</html>
