<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link href="../Css/Manage.Css" rel="stylesheet" type="text/css">
<script language="javascript" src="/Js/All.js"></script>
<script language="javascript" src="/Js/MdiWin.js"></script>
<script language="javascript">
<!--
function ShowLeftBar()
	{
		window.top.Frame_Main.cols="180,*";
    LeftBar.style.display="none";
	}
	
function HideLeftBar()
  {top.Frame_Main.cols="0,*";
   LeftBar.style.display="";
   show=false;
  }

function ReMoveWin(Window)
{
	if (Window == 0)
	{
		if(!confirm("注意：\n \n    是否要关闭所有页面? "))
		{
			return false;
		}
		else
		{
			win.removeall();
		}
	}
	else
	{
			win.removewin(win.currentwin);
	}
}
-->
</script>
<base target="_self">
</head>

<body  onLoad="init()"  scroll="no">
<table border="0" width="100%" height="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td height="21">
			<table height="100%" width="100%" border="0" cellspacing="0" cellpadding="0" style="table-layout: fixed" bgcolor="#D6D3CE">
				<tr height="21" bgcolor="#D6D3CE">
					<td valign="bottom" nowrap width="*">
						<div style="overflow:hidden;width:100%">
							<div id="titlelist" style="margin-left:0;z-index:-1">
							</div>
						</div>
					</td>
					<td align="center" class="Menu" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" width="24" height="15" title="显示菜单" id="LeftBar" onClick="ShowLeftBar()" style="display:none;"><img src="images/Menu_Show.gif"></td>
					<td align="center" class="Menu" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" width="15" height="15" title="向左移动" onmousedown="tabScroll('left')"  onmouseup="tabScrollStop()"><img src="images/Move_Left.gif"></td>
					<td align="center" class="Menu" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" width="15" height="15" title="向右移动" onmousedown="tabScroll('right')" onmouseup="tabScrollStop()"><img src="images/Move_Right.gif"></td>
					<td align="center" class="Menu" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" width="20" height="20" title="关闭一个窗口" onclick="ReMoveWin(1)"><img src="images/Close_One.gif"></td>
					<td align="center" class="Menu" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" width="20" height="20" title="关闭所有窗口" onclick="ReMoveWin(0)"><img src="images/Close_All.gif"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="*">
			<table height="100%" width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td id="mywindows" colspan="5" bgcolor="#D6D3CE" height="100%" width="100%"><p align="center">&nbsp;欢迎使用Monolith Portal 
					后台管理系统，请点击左边菜单进行操作。</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<iframe name="tabWin" marginwidth="1" marginheight="1" scrolling="no" border="0" frameborder="0" src="about:link" width="0" height="0">浏览器不支持嵌入式框架，或被配置为不显示嵌入式框架。</iframe>
</body>
</html>