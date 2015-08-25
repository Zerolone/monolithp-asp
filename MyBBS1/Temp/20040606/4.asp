<!--HTTP头--> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="generator" content="dvbbs">
<meta name=keywords content="8du,8dubbs,8duhost,aspsky,dvbbs,动网,动网论坛,asp,论坛,插件">
<meta name="description" content="八度网络 大型综合论坛。">
<title>八度论坛-论坛首页</title>
<!--默认风格--> 
<style type=text/css>
A:link,A:active,A:visited{TEXT-DECORATION:none ;Color:#000000}A:hover{TEXT-DECORATION: underline;Color:#4455aa}
BODY{FONT-SIZE: 12px;COLOR: #000000;FONT-FAMILY:  宋体;
scrollbar-face-color: #DEE3E7;scrollbar-highlight-color: #FFFFFF;scrollbar-shadow-color: #DEE3E7;scrollbar-3dlight-color: #D1D7DC;scrollbar-arrow-color:  #006699;scrollbar-track-color: #EFEFEF;scrollbar-darkshadow-color: #98AAB1;}
font{line-height : normal ;}
TD{font-family: 宋体;font-size: 12px;line-height : 15px ;}
th{background-image: url(Skins/Default/css/default/bg1.gif);background-color: #4455aa;color: white;font-size: 12px;font-weight:bold;}
td.TableTitle2{background-color: #E4E8EF;}
td.TableBody1{background-color: #FFFFFF;line-height : normal ;}
td.TableBody2{background-color: #E4E8EF;line-height : normal ;}
td.TopDarkNav{background-image: url(Skins/Default/css/default/topbg.gif);}
td.TopLighNav{background-image: url(Skins/Default/css/default/bottombg.gif);}
td.TopLighNav1{background-image: url(Skins/Default/css/default/tabs_m_tile.gif);}
td.TopLighNav2{background-color:#FFFFFF}
.tableBorder1{width:98%;border: 1px; background-color: #6595D6;}
.tableBorder2{width:98%;border: 1px #DEDEDE solid; background-color: #EFEFEF;}
#TableTitleLink A:link, #TableTitleLink A:visited, #TableTitleLink A:active {COLOR: #FFFFFF;	TEXT-DECORATION: none;}#TableTitleLink A:hover {COLOR: #FFFFFF;	TEXT-DECORATION: underline;}
input,select,Textarea,option{font-family:Tahoma,Verdana,"宋体"; font-size: 12px; line-height: 15px;COLOR: #000000;}
.normalTextSmall {     font-size : 11px;    color : #000000;     font-family: Verdana, Arial, Helvetica, sans-serif;}
.menuskin {
	BORDER: #666666 1px solid; VISIBILITY: hidden; FONT: 12px Verdana;
	POSITION: absolute; 
	BACKGROUND-COLOR:#EFEFEF;
	background-image:url("Skins/Default/dvmenubg3.gif");
	background-repeat : repeat-y;
	}
.menuskin A {
	PADDING-RIGHT: 10px; PADDING-LEFT: 25px; COLOR: black; TEXT-DECORATION: none; behavior:url(inc/noline.htc);
	}
#mouseoverstyle {
	BACKGROUND-COLOR: #C9D5E7; margin:2px; padding:0px; border:#597DB5 1px solid;
	}
#mouseoverstyle A {
	COLOR: black
}
.menuitems{
	margin:2px;padding:1px;word-break:keep-all;
}

a.navlink:link {color: #000000; text-decoration:none}
a.navlink:visited {color: #000000; text-decoration:none }
a.navlink:hover {color: #003399; text-decoration:none }
.BrightClass{background-color: #D7D7D7; }
/*
编辑器特效CSS样式
*/
div.quote{margin:5px 20px;border:1px solid #CCCCCC;padding:5px;background:#F3F3F3 ;line-height : normal ;
}
div.HtmlCode{margin:5px 20px;border:1px solid #CCCCCC;padding:5px; background:#FDFDDF ;
font-size:14px;font-family:Tahoma;font-style : oblique;line-height : normal ;font-weight:bold;
}
</style><!--论坛页面开始代码--> 
<script language = "JavaScript" src = "inc/Main.js"></script>
</head>
<body topmargin="0" leftmargin="0">
<div class=menuskin id=popmenu 
      onmouseover="clearhidemenu();highlightmenu(event,'on')" 
      onmouseout="highlightmenu(event,'off');dynamichide(event)" style="Z-index:100"></div>
<!--顶部表格-->
<table cellspacing="0" cellpadding="0" align="center" class=tableborder1>
<tr>
<td width="1"></td>
<td class=TopDarkNav height=9 width="*"></td>
<td width="1"></td>
<tr>
<td width="1" height="70"></td>
<td height="70" class=TopLighNav2>
<TABLE border="0" width="100%" align="center">
<TR><TD align=left width="25%"><a href="http://www.8duhost.com/bbs"><img border=0 src="images/logo.gif"></a></TD><TD Align=center width="65%" id="Top_ads"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="468" height="60"><param name=movie value=""><param name=quality value=high><param name=menu value=false><embed src="http://bbs.dvbbs.net/ad/banner02.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="468" height="60"></embed></object></td>
<td align="right" style="line-height: 15pt" width="10%">
<span style="CURSOR: hand" onClick="window.external.AddFavorite(document.location.href, '八度论坛-论坛首页')" onmousemove="status='收藏本页';" onmouseout="status='';" >收藏本页</span><br><a href="mailto:admin@8duhost.com" target=_blank>联系我们</a><br><a href="boardhelp.asp?boardID=0">论坛帮助</a></td></tr></table>
</td>
<td width="1"></td>
<tr>
<td width="1"></td><td class=TopLighNav height=9 width="*"></td>
<td width="1"></td>
</tr>
<tr>
<td width="1"></td>
<td class=TopLighNav1 height=22 valign="middle" id="Menu">&nbsp;<!--顶部用户导航栏：已登录用户菜单-->
<script language="javascript">
var tmenu='<table width="160" border="0" cellspacing="0" cellpadding="0"><tr><td><!--论坛模板和CSS风格选择部分(含7个部分) --><div class="menuitems"><a href=cookies.asp?action=stylemod&skinid=0&boardid=0>恢复默认设置</a></div><table width="95%" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td>&nbsp;&nbsp;&nbsp;<img id="1" style="cursor:hand" onMouseOver="doClick();" src="Skins/Default/plus.gif" width="15" height="15"><span id="1_" style="cursor: hand"  class="menuitems" ><a  title="您当前正在使用的模板"><font color="#FF0000">默认模板</font></a></span><span id="1_content" style="DISPLAY: none"><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a title="您正在使用的Css风格"><font color="#FF0000">默认风格</font></a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=1">水晶紫色</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=2">秋意盎然</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=3">棕红预览</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=4">紫色淡雅</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=5">雪花飘飘</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=6">ｅ点小镇</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=7">心情灰色</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=8">青青河草</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=9">浓浓绿意</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=10">绿色淡雅</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=11">蓝雅绿</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=12">蓝色庄重</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=13">蓝色水晶</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=14">橘子红了</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=15">红红夜思</a></div><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=1&boardid=0&cssid=16">粉色回忆</a></div></span></td></tr></table><table width="95%" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td>&nbsp;&nbsp;&nbsp;<img id="3" style="cursor:hand" onMouseOver="doClick();" src="Skins/Default/plus.gif" width="15" height="15"><span id="3_" style="cursor: hand"  class="menuitems" >视觉空间</span><span id="3_content" style="DISPLAY: none"><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=3&boardid=0&cssid=0">视觉空间</a></div></span></td></tr></table><table width="95%" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td>&nbsp;&nbsp;&nbsp;<img id="4" style="cursor:hand" onMouseOver="doClick();" src="Skins/Default/plus.gif" width="15" height="15"><span id="4_" style="cursor: hand"  class="menuitems" >淡淡桔香</span><span id="4_content" style="DISPLAY: none"><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=4&boardid=0&cssid=0">淡淡桔香</a></div></span></td></tr></table><table width="95%" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td>&nbsp;&nbsp;&nbsp;<img id="8" style="cursor:hand" onMouseOver="doClick();" src="Skins/Default/plus.gif" width="15" height="15"><span id="8_" style="cursor: hand"  class="menuitems" >可爱女生</span><span id="8_content" style="DISPLAY: none"><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=8&boardid=0&cssid=0">新版女生</a></div></span></td></tr></table><table width="95%" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td>&nbsp;&nbsp;&nbsp;<img id="9" style="cursor:hand" onMouseOver="doClick();" src="Skins/Default/plus.gif" width="15" height="15"><span id="9_" style="cursor: hand"  class="menuitems" >绿色幻想</span><span id="9_content" style="DISPLAY: none"><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=9&boardid=0&cssid=0">绿色幻想</a></div></span></td></tr></table><table width="95%" border="0" cellspacing="0" cellpadding="0" align="center"><tr><td>&nbsp;&nbsp;&nbsp;<img id="10" style="cursor:hand" onMouseOver="doClick();" src="Skins/Default/plus.gif" width="15" height="15"><span id="10_" style="cursor: hand"  class="menuitems" >简约之美</span><span id="10_content" style="DISPLAY: none"><div class="menuitems">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="Skins/Default/minus.gif"><a href="cookies.asp?action=stylemod&skinid=10&boardid=0&cssid=0">简约之美</a></div></span></td></tr></table></td></tr></table>'
</script>
<b>Zerolone</b> <a href="login.asp" >重登录</a> <img border=0 src="Skins/Default/navspacer.gif" align="absmiddle">  <a href="cookies.asp?action=hidden&userid=752">隐身</a> <img border=0 src="Skins/Default/navspacer.gif" align ="absmiddle"> <a onMouseOver="showmenu(event,'<div class=menuitems><a style=font-size:9pt;line-height:14pt; href=JavaScript:openScript(\'messanger.asp?action=new\',500,400)>发短信</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=BoardPermission.asp?boardid=0&action=Myinfo>我能做什么</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=\'query.asp?stype=5&s=2&pSearch=0&nSearch=0\'>我发表的主题</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=\'query.asp?stype=5&s=1&pSearch=0&nSearch=0\'>我参与的主题</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=mymodify.asp>基本资料修改</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=modifypsw.asp>用户密码修改</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=modifyadd.asp>联系资料修改</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=usersms.asp>用户短信服务</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=friendlist.asp>编辑好友列表</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=favlist.asp>用户收藏管理</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=myfile.asp>个人文件管理</a></div>')" style="CURSOR:hand" href="usermanager.asp">用户控制面板</a> 
 <img border=0 src="Skins/Default/navspacer.gif" align ="absmiddle"> 
 <a href="query.asp?boardid=0">搜索</a>
 <img border = 0 src = "Skins/Default/navspacer.gif" align = "absmiddle" >  
   <a  onmouseover=showmenu(event,tmenu) onmouseout=delayhidemenu() class='navlink' style="CURSOR:hand" >风格</a> <img border = 0 src = "Skins/Default/navspacer.gif" align = "absmiddle" >
   <a  onMouseOver="showmenu(event,'<div class=menuitems><a style=font-size:9pt;line-height:14pt; href=boardstat.asp?boardid=0 >今日贴数图例</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=boardstat.asp?action=lasttopicnum&boardid=0>主题数图例</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=boardstat.asp?action=lastbbsnum&boardid=0>总帖数图例</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=boardstat.asp?reaction=online&boardid=0>在线图例</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=boardstat.asp?reaction=onlineinfo&boardid=0>在线情况</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=boardstat.asp?reaction=onlineUserinfo&boardid=0>用户组在线图例</a></div>')" style="CURSOR:hand" >论坛状态</a>    <img border = 0 src = "Skins/Default/navspacer.gif" align = "absmiddle" >   <a href="show.asp?boardid=0" onMouseOver="showmenu(event,' <div class=menuitems><a style=font-size:9pt;line-height:14pt; href=show.asp?filetype=0&boardid=0>文件集浏览</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=show.asp?filetype=1&boardid=0>图片集浏览</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=show.asp?filetype=2&boardid=0>Flash浏览</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=show.asp?filetype=3&boardid=0>音乐集浏览</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=show.asp?filetype=4&boardid=0>电影集浏览</a></div><div class=menuitems><a style=font-size:9pt;line-height:14pt; href=show.asp>贺卡发送</a></div>')">论坛展区</a>  <img src=Skins/Default/navspacer.gif align=absmiddle> <a href="#" title="" onMouseOver="showmenu(event,'<div class=menuitems><a href=http://www.8duhost.com/bbs/plus/wish/wish.asp title= target=_blank>许愿瓶</a></div><div class=menuitems><a href=http://www.8duhost.com/bbs/plus_bank.asp title=>社区银行</a></div><div class=menuitems><a href=http://www.8duhost.com/bbs/plus_bz_gongzi.asp title=>版主评定</a></div><div class=menuitems><a href=http://www.8duhost.com/bbs/story.asp title=>故事接龙</a></div><div class=menuitems><a href=http://www.8duhost.com/bbs/prison.asp title=>社区监狱</a></div>')"><font color=red>论坛设施</font></a> <img src="skins/default/navspacer.gif" align=absmiddle> <a href="admin_index.asp">管理</a> <img src="skins/default/navspacer.gif" align=absmiddle> <a href="recycle.asp">回收站</a>  <img border = 0 src = "Skins/Default/navspacer.gif" align = "absmiddle" >   <a href="logout.asp">退出</a> </td>
<td width="1"></td>
</tr>
</table><!--index.asp##首页顶部表格-->
<br><table cellspacing=1 cellpadding=3 align=center border=0 width=98%><tr><td align=center width=100% valign=middle colspan=2>
<a href="javascript:openScript('announcements.asp?action=showone&boardid=0',500,400)"><B>八度需要大家。请大家多多发帖~~~~~~~~~</B></a>(2004-6-1 17:40:29)
</td></tr></table>
<table cellspacing=1 cellpadding=3 align=center class=tableBorder1><tr>
<th width="100%" height="24" colspan="2">
</th></tr>
<tr><td class=tableBody1 align=center width="*" >
<!--index.asp##首页顶部表格已登录-->
<TABLE border=0 width="98%" align=center>
<tr>
<td rowspan=6 width=88 align=center><img src=http://www.digichina.cn/club/Images/userface/image334.gif width=60 height=60></td>
</tr>
<tr><td height=22>
<img src="skins/default/msg_no_new_bar.gif" align="absmiddle"> <a href="usersms.asp?action=inbox">我的收件箱</a> (<font color=gray>0</font>)
</td><td colspan=2>

</td></tr>
<tr><td height="1" colspan="3">
<table border=0 cellspacing=0 cellpadding=0 align=center style="width:100%" class=tableborder1>
<tr><td> </td></tr>
</table>
</td></tr>
<tr><td>
&nbsp;:: <a href="friendlist.asp">我的外交状况</a>
</td><td>
&nbsp;:: <A HREF="javascript:openScript('messanger.asp?action=new',500,400)">发送论坛短信</A>
</td></tr><tr><td>
&nbsp;:: <a href="query.asp?stype=5&s=2&pSearch=0&nSearch=0">我发起的主题</a>
</td><td>
&nbsp;:: <a href="BoardPermission.asp?boardid=0&action=Myinfo">我的论坛权限</a>
</td></tr><tr><td>
&nbsp;:: <a href="query.asp?stype=5&s=1&pSearch=0&nSearch=0">我参与的主题</a>
</td><td>
&nbsp;:: <A HREF="favlist.asp">我的论坛收藏</A>
</td></tr></table>
</td><td width=340 class=tableBody1 valign="top">
<TABLE border=0 width="98%" align=center><tr><td height=24>
共有 <B>1692</B> 位会员
</td><td>
<a href="toplist.asp?orders=2">新进来宾</a> [<a href=dispuser.asp?name=天许骆严 target=_blank><b>天许骆严</b></a>]</td>
</tr><tr><td height="1" colspan="2">
<table border=0 cellspacing=0 cellpadding=0 align=center style="width:100%" class=tableborder1>
<tr><td> </td></tr>
</table>
</td></tr><TR><TD>
今日发帖：<FONT COLOR=#FF0000><B>148</B></FONT> 篇</td><TD>
主题总数：<b>5906</b> 篇
</td></tr><tr><td>
昨日发帖：<B>125</B> 篇
</td><td>
帖子总数：<b>67195</b> 篇
</td></tr>
<tr><td colspan="2">
最高日发帖：<font title="发生日期: [2003-7-13 16:36:00"><B>1645</B></font> 篇，发生时间：2003-7-13 16:36:00
</td></tr></table>
</td>
</tr><tr><td class=tableBody2 align="center" colspan="2" height="25" width="100%">
<a href=query.asp?stype=3&pSearch=0&nSearch=0>查看新贴</a> <font face=Wingdings color=666666>v</font> <a href=query.asp?stype=4&pSearch=0&nSearch=0>热门话题</a> <font face=Wingdings color=666666>v</font> <a href=toplist.asp?orders=1>发贴排行</a> <font face=Wingdings color=666666>v</font> <a href=toplist.asp?orders=7>用户列表</a></td></tr>
</table><br><!--index.asp##处理版面循环列表JS部分-->
<script language="javascript">
var template=new Array();
var piclist=new Array();
var Strings=new Array();
var mainsetting=new Array();
var ShowMasters,tablecount,tablewidth,islist,Divstr1,Divstr2,listst,RootID=0,Child=0,boardcount=0;
var tmpstr1,tmpstr2
var k
//显示分类头表格，入口参数BoardID,BoardType,Board_Setting
function showclass(BoardID,BoardType,Board_Setting,Cookies,vChild)
{
	Board_Setting=Board_Setting.split(",")
	var pic='Skins/Default/nofollow.gif';
	var thetitle=Strings[3]
	tablecount=(Board_Setting[41]);
	tablewidth=Math.floor(100/tablecount);
	tablewidth+="%"
	var Divstr1=''
	var Divstr2='style=display:none'
	RootID=BoardID
	Child=(vChild)
	boardcount=0
	var str=template[0]
	if(Cookies=='')
	{
		islist=(Board_Setting[39])
	}
	else
	{
		islist=(Cookies)
	}
	if (islist!="0")
	{
		pic='Skins/Default/plus.gif';
		thetitle=Strings[4]
		Divstr2=''
		Divstr1='style=display:none'
		str = str.replace(/{\$value}/gi,"0");
	}
	else
	{
		str = str.replace(/{\$value}/gi,"1");
	}
	
	str = str.replace(/{\$boardid}/gi,BoardID);
	str = str.replace(/{\$pic}/gi,pic);
	str = str.replace(/{\$title}/gi,thetitle);
	str = str.replace(/{\$BoardType}/gi,BoardType);
	str = str.replace(/{\$Divstr1}/gi,Divstr1);
	str = str.replace(/{\$Divstr2}/gi,Divstr2);
	document.write (str);
	k=0
}

function showboard(BoardID,BoardType,child,readme,BoardMaster,PostNum,TopicNum,indexIMG,todayNum,LastPost,Board_Setting,havenew)
{	
	boardcount++
	k++
	var str=template[2]
	var str1=template[4]
	Board_Setting=Board_Setting.split(",")
	BoardMaster=BoardMaster.split("|")
	ShowMasters=''
	ShowMaster=''
	var tmp
	for(j=0;j<BoardMaster.length;j++)
	{
		if (j>5){ShowMasters+=' <font color=gray>More</font>';break;}
		ShowMasters+='&nbsp;<a  onMouseOver=\'showmenu(event,"<a style=font-size:9pt;line-height:14pt; href=dispuser.asp?name='+BoardMaster[j]+' target=_blank>资料</a><br><a style=font-size:9pt;line-height:14pt; href=messanger.asp?action=new&touser='+BoardMaster[j]+' target=_blank>留言</a>")\'>'+BoardMaster[j]+'</a>';
	}
	for(j=0;j<BoardMaster.length;j++)
	{
		if (j>1){ShowMaster+=' <font color=gray>More</font>';break;}
		ShowMaster+='&nbsp;<a  onMouseOver=\'showmenu(event,"<a style=font-size:9pt;line-height:14pt; href=dispuser.asp?name='+BoardMaster[j]+' target=_blank>资料</a><br><a style=font-size:9pt;line-height:14pt; href=messanger.asp?action=new&touser='+BoardMaster[j]+' target=_blank>留言</a>")\'>'+BoardMaster[j]+'</a>';
	}
	str = str.replace(/{\$boardid}/gi,BoardID);
	str1 = str1.replace(/{\$boardid}/gi,BoardID);
	str1 = str1.replace(/{\$width}/gi,tablewidth);
	str1 = str1.replace(/{\$blinkcolor}/gi,mainsetting[3]);
	str1 = str1.replace(/{\$alertcolor}/gi,mainsetting[1]);
	str1 = str1.replace(/{\$ShowMasters}/gi,ShowMaster);
	str1 = str1.replace(/{\$todayNum}/gi,todayNum);
	str1 = str1.replace(/{\$PostNum}/gi,PostNum);
	str1 = str1.replace(/{\$TopicNum}/gi,TopicNum);
	str1 = str1.replace(/{\$readme}/gi,HTMLEncode(readme));
	str = str.replace(/{\$BoardType}/gi,BoardType);
	str1 = str1.replace(/{\$BoardType}/gi,BoardType);
	
	if (child==0)
	{
		str = str.replace(/{\$child}/gi,'');
		str1 = str1.replace(/{\$child}/gi,'');
	}
	else
	{
		tmp = Strings[1];
		tmp = tmp.replace(/{\$child}/gi,child);
		str1 = str1.replace(/{\$child}/gi,tmp);
		str = str.replace(/{\$child}/gi,tmp);
		Child=Child-child;		
	}
	str = str.replace(/{\$statuspic}/gi,showpic(havenew,Board_Setting[0],Board_Setting[2]));
	str = str.replace(/{\$readme}/gi,readme);
	str = str.replace(/{\$ShowMasters}/gi,ShowMasters);
	str = str.replace(/{\$PostNum}/gi,PostNum);
	str = str.replace(/{\$TopicNum}/gi,TopicNum);
	str = str.replace(/{\$todayNum}/gi,todayNum);
	str = str.replace(/{\$blinkcolor}/gi,mainsetting[3]);
	str = str.replace(/{\$alertcolor}/gi,mainsetting[1]);
	if (Board_Setting[2]=='1')
	{
		str = str.replace(/{\$LastPost}/gi,Strings[2]);
	}
	else
	{
		str = str.replace(/{\$LastPost}/gi,showlastpost(LastPost));
	}
	
	if (indexIMG!='')
	{
		str = str.replace(/{\$indexIMG}/gi,'<table align="left"><tr><td><a href="list.asp?boardid='+BoardID+'"><img src='+indexIMG+' align="top" border="0"></a></td><td width="20"></td></tr></table>');
	}
	else
	{
		str = str.replace(/{\$indexIMG}/gi,'');
	}
	if(k==tablecount)
	{
		str1+="</tr>"
		if (boardcount!=Child)
		{str1+="<tr>"}
		k=0
	}
	
	showcode(str,str1)
}
function showcode(str,str1)
{	
	if (boardcount==1)
	{
		tmpstr1=template[1]
		tmpstr2=template[1]
		tmpstr2+="<tr>"
	}
	tmpstr1+=str
	tmpstr2+=str1
}
function classfooter()
{
	if (k!=0)
		{	template[5]=template[5].replace(/{\$width}/gi,tablewidth);
			for(jj=0;jj<tablecount-k;jj++)
			{
			tmpstr2+=template[5]
			}
		}
		tmpstr1+=template[3]
		tmpstr2+="</tr>"
		tmpstr2+=template[3]
		document.getElementById("ListDiv1_"+RootID).innerHTML=tmpstr1
		document.getElementById("ListDiv2_"+RootID).innerHTML=tmpstr2
		tmpstr1='';
		tmpstr2='';
		boardcount=0
}

function showpic(havenew,Board_Setting,Board_Setting1)
{
	var pic,Str,Str1
	Str="无新贴"
	Str1="开放的版面"
	pic=piclist[0]
	if(havenew==1)
	{
		Str="有新贴"
		pic=piclist[1]
	}
	if(Board_Setting==1) 
	{
		pic=piclist[2]
		Str1="锁定的版面"
	}
	if(Board_Setting1==1) 
	{
		pic=piclist[2]
		Str1="认证论坛"
	}		
	return('<img src="'+pic+'" alt="'+Str1+','+Str+'">')
} 
function showlastpost(lastpoststr)
{
	lastpoststr=lastpoststr.replace(/</gi,'&lt;');
	if (lastpoststr=='$$$$'||lastpoststr=='')
	{
		return('主题：无<br>作者：无<br>日期：无')
	}
	else
	{
		var str='';
		lastpoststr=lastpoststr.split("$");
		str+='主题：<a href="Dispbbs.asp?boardid='+lastpoststr[7]+'&ID='+lastpoststr[6]+'&replyID='+lastpoststr[1]+'&skin=1" title="转到：'+lastpoststr[3]+'">';
		str+=lastpoststr[3].substring(0,10);
		str+='</a>';
		str+='<br>作者：';
		str+='<a href="dispuser.asp?id='+lastpoststr[5]+'" target="_blank">'+lastpoststr[0] +'</a>';
		str+='<br>日期：';
		str+=lastpoststr[2] +'&nbsp;<a href="dispbbs.asp?Boardid='+lastpoststr[7]+'&ID='+ lastpoststr[6] +'&replyID=' + lastpoststr[1] +'&skin=1"><IMG border=0 src="Skins/Default/lastpost.gif" title="主题：'+lastpoststr[3]+'"></a>';
		return(str)
	}

}
function ReShowList(ListID)
{
	var ListDiv1=document.getElementById("ListDiv1_"+ListID);
	var ListDiv2=document.getElementById("ListDiv2_"+ListID);
	var titlepic=document.getElementById("titlepic_"+ListID);
	if (ListDiv1.style.display=='none')
	{
		ListDiv2.style.display='none';
		ListDiv1.style.display='block';
		titlepic.innerHTML='<a href="cookies.asp?action=setlistmod&thisvalue=1&id='+ListID+'" target="hiddenframe"><img src="Skins/Default/nofollow.gif" border=0 title='+Strings[3]+'></a>';
		
	}
	else
	{
		ListDiv1.style.display='none';
		ListDiv2.style.display='block';
		titlepic.innerHTML='<a href="cookies.asp?action=setlistmod&thisvalue=0&id='+ListID+'" target="hiddenframe"><img src="Skins/Default/plus.gif" border=0 title='+Strings[4]+'></a>';
		
	}
}
</script><script language="javascript">
piclist[0]='Skins/Default/forum_nonews.gif';
piclist[1]='Skins/Default/forum_isnews.gif';
piclist[2]='Skins/Default/forum_lock.gif';
piclist[3]='Skins/Default/birthday.gif';
mainsetting[0]='98%';
mainsetting[1]='#FF0000';
mainsetting[2]='#00008b';
mainsetting[3]='#000066';
mainsetting[4]='#000066';
mainsetting[5]='#9898BA';
mainsetting[6]='#990000';
mainsetting[7]='#9898BA';
mainsetting[8]='#990000';
mainsetting[9]='#9898BA';
mainsetting[10]='#0000ff';
mainsetting[11]='#cccccc';
mainsetting[12]='#6595D6';
template[template.length]='<!--index.asp##首页一级分类表格-->\n<table cellspacing=1 cellpadding=0 align=center class=tableBorder1>\n  <TR>\n    <Th colSpan=2 height=25 align=left id=TableTitleLink>&nbsp;<span id="titlepic_{$boardid}"><a href="cookies.asp?action=setlistmod&thisvalue={$value}&id={$boardid}" target="hiddenframe"><img src="{$pic}" border=0 title={$title}></a></span> <a href="list.asp?boardid={$boardid}" title=进入本分类论坛>{$BoardType}</a><Th>\n  </tr>\n</table>\n<Div id=ListDiv1_{$boardid} {$Divstr1}><!--列表层 --></div>\n<Div id=ListDiv2_{$boardid} {$Divstr2}><!--简洁层 --></div>\n<br>';
template[template.length]='<!--index.asp##处理版面循环列表头尾table部分合集-->\n<table cellspacing=1 cellpadding=0 align=center class=tableBorder1>\n';
template[template.length]='<!--index.asp##首页版面循环（列表）-->\n<tr>\n    <td valign="middle" width="46" align="center" class="tablebody1"> \n      {$statuspic}\n    </td>\n<TD vAlign="top" width="*" class="tablebody1">\n<table cellSpacing="0" cellPadding="2" width="100%" border="0">\n<tr>\n          <td class="tablebody1" width="*" height=20 > <a href="list.asp?boardid={$boardid}"><font color={$blinkcolor}>{$BoardType}</font></a>{$child}\n          </td>\n          <td width="40" rowspan="2" align="center" class="tablebody1"> \n           {$indexIMG}\n          </td>\n          <td width="200" rowspan="2" class="tablebody1">{$LastPost}\n          </td>\n</tr>\n<tr>\n          <TD width=*>\n            <FONT face=Arial><img src="Skins/Default/Forum_readme.gif" align=middle>{$readme}</font> </td>\n</tr>\n<TR><TD class=tablebody2 height=20 width=*>版主：{$ShowMasters}\n</td>\n<td width="40" align="center" class="tablebody2">&nbsp;</td><TD valign="middle" class="tablebody2" width="200">\n<table width="100%" border="0">\n<tr>\n<td width="25%" valign="middle"><img src="Skins/Default/Forum_today.gif"  alt="今日发贴数" align="absmiddle">&nbsp;<font color="{$alertcolor}">{$todayNum}</font></td>\n                <td width="30%" valign="middle"><img src="Skins/Default/forum_topic.gif" alt="主题贴数" border="0"  align="absmiddle">&nbsp;{$TopicNum}\n                </td>\n                <td width="45%" vAlign="middle">\n<img src="Skins/Default/Forum_post.gif" alt="发贴总数" border="0" align="absmiddle" >&nbsp;{$PostNum}</td>\n</tr>\n</table>\n</td>\n</tr>\n</table>\n</td>\n</tr>';
template[template.length]='\n</table>\n';
template[template.length]='<!--index.asp##首页版面循环（简洁）-->\n<td class=tablebody1 width={$width}>\n<TABLE cellSpacing=2 cellPadding=2 width=100% border=0>\n<tr>\n<td width="100%" title="{$readme}" colspan=3>\n<a href=list.asp?boardid={$boardid}><font color={$blinkcolor}>{$BoardType}</font></a>{$child}</td></tr>\n<tr>\n<td width="25%" valign="middle"><img src="Skins/Default/Forum_today.gif"  alt="今日发贴数" align="absmiddle">&nbsp;<font color="{$alertcolor}">{$todayNum}</font></td>\n                <td width="30%" valign="middle"><img src="Skins/Default/forum_topic.gif" alt="主题贴数" border="0"  align="absmiddle">&nbsp;{$TopicNum}\n                </td>\n                <td width="45%" vAlign="middle">\n<img src="Skins/Default/Forum_post.gif" alt="发贴总数" border="0" align="absmiddle" >&nbsp;{$PostNum}</td>\n</tr>\n<TR><TD class=tablebody2 height=20 width=* colspan=3>版主：{$ShowMasters}\n</td></tr>\n</table></td>';
template[template.length]='\n<td class=tablebody1 width={$width}>\n<TABLE cellSpacing=2 cellPadding=2 width=100% border=0>\n<tr>\n    <td width="100%" title="" colspan=3>&nbsp; </td>\n  </tr>\n<tr>\n    <td width="25%" valign="middle">&nbsp;</td>\n                \n    <td width="30%" valign="middle">&nbsp;</td>\n                \n    <td width="45%" vAlign="middle">&nbsp; </td>\n</tr>\n<TR>\n    <TD class=tablebody2 height=20 width=* colspan=3>&nbsp;</td>\n  </tr>\n</table></td>';
template[template.length]='<!--index.asp##首页友情连接循环部分-->\n<td width="16%">&nbsp;{$link}</td>';
Strings[Strings.length]='论坛首页';Strings[Strings.length]='（<font title=有{$child}个下属论坛>{$child}</font>）';Strings[Strings.length]='认证论坛，请认证用户进入浏览';Strings[Strings.length]='切换到简洁模式';Strings[Strings.length]='切换到列表模式';Strings[Strings.length]='尚无友情论坛';Strings[Strings.length]='显示详细列表';Strings[Strings.length]='关闭详细列表';Strings[Strings.length]='当前没有公告';Strings[Strings.length]='今天没有朋友生日';Strings[Strings.length]='&nbsp;<a href=dispuser.asp?name={$username} title="祝{$age}岁生日快乐！" target="_blank">〖祝{$username} 生日快乐<img src="{$bpic}" align=absmiddle border=0 width="20" height="20">〗</a>';showclass(5,'..::::.╃ 八度网络综合版╃ .:::..','0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,1,0,30,20,10,9,12,1,10,10,0,0,0,0,1,4,0,1,4,0,0,0,100,0,0,,$$,0,0,0,0,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0','',9);
showboard(12,'┫&#9618; 灌水乐园&#9618;┣ ',0,'<font color=\\\"black\\\">灌水天堂，不管你有什么，都可以在这里畅谈。 想说什么，就说什么。。。。。。严禁纯水！</font><font color=\\\"red\\\"></font>','安静的苦笑|o0乖怪0o ',35187,2731,'',12,'lllllll$67571$2004-6-6 21:33:47$请进$jpg$2256$6634$12','0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,1,0,30,20,10,9,12,1,10,10,0,0,0,0,1,4,0,1,4,0,0,0,100,0,0,,$$,0,0,0,0,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0',1);
showboard(94,'┣&#9618; 情感随笔&#9618;┣ ',0,' ','阿Q|lulu|尘封の绝恋',12034,743,'',117,'lulu$67570$2004-6-6 20:43:51$答应陪你,真的$$433$6632$94','0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,1,0,30,20,10,9,12,1,10,10,0,0,0,0,1,3,0,1,4,0,0,0,100,0,0,,$$,0,0,0,0,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0',0);
showboard(6,'┫&#9618; 卡通，漫画 &#9618;┣ ',1,'<font color=\\\\\\\"black\\\\\\\">最流行的卡通漫画的讨论版块...激情无限卡通无限！</font>','',7357,578,'',16,'lllllll$67572$2004-6-6 21:38:19$AFGSDFH$jpg$2256$6635$6','0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,1,0,30,20,10,9,12,1,5,10,0,0,0,0,1,4,0,1,4,0,0,0,100,0,0,,$$,0,0,0,0,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,1440,0,0,0,0,0,0,0,0,0',1);

showboard(7,'┫&#9618; Top 流行 &#9618;┣ ',2,'<font color=\\\"black\\\">音乐，街舞，时尚的空间时尚的版块！还等什么？快点加入我们吧~</font>','星矢|niaijiang',1093,332,'',0,'lulu$67422$2004-6-5 22:35:37$有喜欢twins的么?!$$433$5598$7','0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,1,0,30,20,10,9,12,1,10,10,0,0,0,0,1,3,0,1,4,0,0,0,100,0,0,,$$,0,0,0,0,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0',0);


showboard(105,'┫&#9618; 网游技术专区 &#9618;┣  ',0,'<font color=\\\"black\\\">网络游戏讨论基地</font>','o0乖怪0o|天行者',2471,387,'',1,'泡泡$67546$2004-6-6 17:21:33$八度工会&分工会成员来...$$1864$2280$105','0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,1,0,30,20,10,9,12,1,10,10,0,0,0,0,1,3,0,1,4,0,0,0,100,0,0,,$$,0,0,0,0,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0',0);
showboard(107,'┫&#9618; 星座情缘&#9618;┣ ',0,'<font color=\\\"black\\\">12宫轮回运转，命运随之起起伏伏。心系命运，关注星座神话…</font>','雨雪琪霏',2141,412,'',0,'蓝色幻想$67292$2004-6-4 22:13:13$看看你和哪个明星一天...$$2114$6616$107','0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,1,0,30,20,10,9,12,1,10,10,0,0,0,0,1,3,0,1,4,0,0,0,100,0,0,,$$,0,0,0,0,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0',0);
classfooter();showclass(95,'..::::.╃ 八度网络休闲--个性版╃ .:::..','0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,1,0,30,20,10,9,12,1,10,10,0,0,0,0,1,3,0,1,4,0,0,0,100,0,0,,$$,0,0,0,0,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0','',2);
showboard(96,'异族部落',0,'<font color=\\\"black\\\">你喜歡冒險,想享受刺激嗎?所有世上的奇闻怪事,在这里你都能看到...........異鏃蔀落能給你帶來冒險與刺激的綜合感覺,100%酷到極點!!\\n\\r</font>','wendy',3838,263,'',1,'banya$67423$2004-6-6 0:16:54$会捕鱼的蜘蛛$jpg$1$6621$96','0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,1,0,30,20,10,9,12,1,10,10,0,0,0,0,1,3,0,1,4,0,0,0,100,0,0,,$$,0,0,0,0,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0',0);
showboard(97,'→『失落灵魂收容所』 ',0,' <font color=\\\"black\\\">愿破碎失落的灵魂在这个收容所里早日康复.</font>','wendy|安静的苦笑',464,47,'',1,'安静的苦笑$67506$2004-6-6 13:59:25$高考的目的$$1459$6631$97','0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,1,0,30,20,10,9,12,1,10,10,0,0,0,0,1,3,0,1,4,0,0,0,100,0,0,,$$,0,0,0,0,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0',0);
classfooter();showclass(3,'..::::.╃ 八度网络----社区站务╃ .:::..','0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,1,0,30,20,10,9,12,1,10,10,0,0,0,0,1,4,0,1,4,0,0,0,100,0,0,论坛升级专题$$,$$,0,0,0,0,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0','',2);
showboard(53,'〖站务处理〗',0,'本站站务公告和通知，任命斑竹，或者解除斑竹的职务，都在这里公开发布。','o0小宝0o',1890,164,'',1,'lllllll$67425$2004-6-6 1:23:47$呵呵!$jpg$2256$6628$53','0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,1,0,30,20,10,9,12,1,10,10,0,0,0,0,1,3,0,1,4,0,0,0,100,0,0,,$$,0,0,0,0,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0',0);
showboard(92,'『管理员办公室』',0,'论坛所有管理……<br>欢迎各位积极提出意见和建议。','banya|o0小宝0o|尘封の绝恋',506,56,'',0,'蓝色幻想$66655$2004-5-23 13:05:56$可是试用期还没...$$2114$6303$92','0,0,0,0,0,0,1,1,1,1,1,1,1,1,1,1,16240,3,0,gif|jpg|jpeg|bmp|png|rar|txt|zip|mid,0,0,1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1,1,0,30,20,10,9,12,1,10,10,0,0,0,0,1,3,0,1,4,0,0,0,100,0,0,,$$,0,0,0,0,0|0|0|0|0|0|0|0|0,0|0|0|0|0|0|0|0|0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0',0);
classfooter();
</script><BR><table cellpadding=6 cellspacing=1 align=center class=tableBorder1><TR><Th colSpan=4 align=left>-=> 社区明星&nbsp;[<a href=index.asp?action=&action1=mingxing>我不想看到他们</a>]</th></tr><tr><TD vAlign=middle align=middle width='5%' class=tablebody2> <IMG height=27 src='mingxing/Male.gif' width=25 align=absMiddle></TD><TD vAlign=top align=middle width='45%' class=tablebody1><table border=0 width='100%'><tr><td width=44% valign=top> <strong><font color=#ff0000>今日</font>发帖<font color=#DA9136>状元</font></strong> <font color=#9999ff>照片-=></font><p align=left>姓　　名：<font color=blue>xlyxiaoiu</font> <br>社区等级：论坛游民<br><font color=red>今日</font>发贴：<font color=red> 38</font> 篇<br><table cellspacing=0 cellpadding=0 width=150 border=0><tr><td>个人财产：</td><td><table cellspacing=0 cellpadding=0 width=86 border=0><tbody><tr><td width=3 height=13 ><img height=13 src=mingxing/img_left.gif width=3></td><td align=left width=80 background=mingxing/img_backing.gif height=13 ><img src=mingxing/bar4.gif width=11.42 height=9 alt='1284元现金'><IMG height=9 src=mingxing/bar4_1.gif border=0></td><td width=3 height=13 ><img height=13 src=mingxing/img_right.gif width=3></td></tr></tbody></table></td></tr></table><table cellspacing=0 cellpadding=0 width=150 border=0><tr><td>个人魅力：</td><td><table cellspacing=0 cellpadding=0 width=86 border=0><tbody><tr><td width=3 height=13 ><img height=13 src=mingxing/img_left.gif width=3></td><td align=left width=80 background=mingxing/img_backing.gif height=13 ><img src=mingxing/bar5.gif width=5.88 height=9 alt=176><IMG height=9 src=mingxing/bar5_1.gif border=0></td><td width=3 height=13 ><img height=13 src=mingxing/img_right.gif width=3></td></tr></tbody></table></td></tr></table><table cellspacing=0 cellpadding=0 width=150 border=0><tr><td>社区经验：</td><td><table cellspacing=0 cellpadding=0 width=86 border=0><tbody><tr><td width=3 height=13 ><img height=13 src=mingxing/img_left.gif width=3></td><td align=left width=80 background=mingxing/img_backing.gif height=13 ><img src=mingxing/bar3.gif width=7 height=9 alt=400><IMG height=9 src=mingxing/bar3_1.gif border=0></td><td width=3 height=13 ><img height=13 src=mingxing/img_right.gif width=3></td></tr></tbody></table></td></tr></table>E&nbsp;-&nbsp;mail：<a href='mailto:youxiaoiu@etang.com'><font color=olive>飞鸽传书</font></a></td><td width=50% valign=middle align=center> <a href=javascript:openScript('dispuser.asp?name=xlyxiaoiu',800,500)><img border=0 src='uploadFace/2246_2004611551020545.gif' width='120' height='120'alt='看什么看，偶是八度论坛今日明星！有本事将俺赶走啊！我等着你哦！呵呵！'></a></td></table></TD><TD vAlign=middle align=middle width='5%' class=tablebody2> <IMG height=27 src='mingxing/FeMale.gif' width=25 align=absMiddle></TD><TD vAlign=top align=middle width='45%' class=tablebody1><table border=0 width='100%'><tr><td width=44% valign=top> <strong><font color=#ff0000>今日</font>发帖<font color=#DA9136>榜眼</font></strong> <font color=#9999ff>照片-=></font><p align=left>姓　　名：<font color=blue>lulu</font> <br>社区等级：超级版主<br><font color=red>今日</font>发贴：<font color=red> 26</font> 篇<br><table cellspacing=0 cellpadding=0 width=150 border=0><tr><td>个人财产：</td><td><table cellspacing=0 cellpadding=0 width=86 border=0><tbody><tr><td width=3 height=13 ><img height=13 src=mingxing/img_left.gif width=3></td><td align=left width=80 background=mingxing/img_backing.gif height=13 ><img src=mingxing/bar4.gif width=37.46 height=9 alt='6492元现金'><IMG height=9 src=mingxing/bar4_1.gif border=0></td><td width=3 height=13 ><img height=13 src=mingxing/img_right.gif width=3></td></tr></tbody></table></td></tr></table><table cellspacing=0 cellpadding=0 width=150 border=0><tr><td>个人魅力：</td><td><table cellspacing=0 cellpadding=0 width=86 border=0><tbody><tr><td width=3 height=13 ><img height=13 src=mingxing/img_left.gif width=3></td><td align=left width=80 background=mingxing/img_backing.gif height=13 ><img src=mingxing/bar5.gif width=35.165 height=9 alt=6033><IMG height=9 src=mingxing/bar5_1.gif border=0></td><td width=3 height=13 ><img height=13 src=mingxing/img_right.gif width=3></td></tr></tbody></table></td></tr></table><table cellspacing=0 cellpadding=0 width=150 border=0><tr><td>社区经验：</td><td><table cellspacing=0 cellpadding=0 width=86 border=0><tbody><tr><td width=3 height=13 ><img height=13 src=mingxing/img_left.gif width=3></td><td align=left width=80 background=mingxing/img_backing.gif height=13 ><img src=mingxing/bar3.gif width=36.345 height=9 alt=6269><IMG height=9 src=mingxing/bar3_1.gif border=0></td><td width=3 height=13 ><img height=13 src=mingxing/img_right.gif width=3></td></tr></tbody></table></td></tr></table>E&nbsp;-&nbsp;mail：<a href='mailto: wendy_860208@163.com'><font color=olive>飞鸽传书</font></a></td><td width=50% valign=middle align=center> <a href=javascript:openScript('dispuser.asp?name=lulu',800,500)><img border=0 src='http://www.8duhost.com/bbs/UploadFile/20043302191326229.bmp' width='120' height='120'alt='看什么看，偶是八度论坛今日明星！有本事将俺赶走啊！我等着你哦！呵呵！'></a></td></table></TD></TR><tr><td colSpan=4 class=tablebody2><font color=#ff0000><b>10TOP</b></font><font color=#1E90FF>&nbsp;<font face=Wingdings color=blue>v</font> <a href=dispuser.asp?name=一乘寺贤 target=_blank title="查看该用户个人资料及最新发帖一览"><font color=#00BFFF>一乘寺贤</font></a>[<font color=red>3674</font>]&nbsp;&nbsp;<font face=Wingdings color=blue>v</font> <a href=dispuser.asp?name=lulu target=_blank title="查看该用户个人资料及最新发帖一览"><font color=#00BFFF>lulu</font></a>[<font color=red>2933</font>]&nbsp;&nbsp;<font face=Wingdings color=blue>v</font> <a href=dispuser.asp?name=wendy target=_blank title="查看该用户个人资料及最新发帖一览"><font color=#00BFFF>wendy</font></a>[<font color=red>2806</font>]&nbsp;&nbsp;<font face=Wingdings color=blue>v</font> <a href=dispuser.asp?name=o0小宝0o target=_blank title="查看该用户个人资料及最新发帖一览"><font color=#00BFFF>o0小宝0o</font></a>[<font color=red>2627</font>]&nbsp;&nbsp;<font face=Wingdings color=blue>v</font> <a href=dispuser.asp?name=o0乖怪0o target=_blank title="查看该用户个人资料及最新发帖一览"><font color=#00BFFF>o0乖怪0o</font></a>[<font color=red>2270</font>]&nbsp;&nbsp;<font face=Wingdings color=blue>v</font> <a href=dispuser.asp?name=丶.桭． target=_blank title="查看该用户个人资料及最新发帖一览"><font color=#00BFFF>丶.桭．</font></a>[<font color=red>2171</font>]&nbsp;&nbsp;<font face=Wingdings color=blue>v</font> <a href=dispuser.asp?name=林子 target=_blank title="查看该用户个人资料及最新发帖一览"><font color=#00BFFF>林子</font></a>[<font color=red>2164</font>]&nbsp;&nbsp;<font face=Wingdings color=blue>v</font> <a href=dispuser.asp?name=x target=_blank title="查看该用户个人资料及最新发帖一览"><font color=#00BFFF>x</font></a>[<font color=red>2118</font>]&nbsp;&nbsp;</td></tr></font></table><br>
<!--index.asp##首页友情连接HTML+JS处理部分-->
<BR>
<!--index.asp##首页今日生日用户部分-->
<br>
<!--index.asp##首页用户信息和在线部分和底部论坛状态-->
<BR>
<!--index.asp##首页底部论坛状态部分-->
<iframe width="0" height="0" src="Online.asp?action=1&Boardid=0" name="hiddenframe"></iframe><script language="javascript">

</script><!--页面结束部分--> 
<Div id=bottom align="center" ></div>
<br>
</body></html>