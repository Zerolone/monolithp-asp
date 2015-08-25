<%
'---------------------------首页主板块的默板：需要替换的参数为, BoardName, BoardID
 dim Skin_BoardMain		
  Skin_BoardMain = "<table border='0' width='100%'>"
  Skin_BoardMain = Skin_BoardMain + "<tr><td width='100%' colspan='5' background='images/tile_back.gif' height='36'>　<img src='images/nav_m.gif' border='0'>&nbsp;<font color='#FFFFFF'>{BoardName}</font></td></tr>"
  Skin_BoardMain = Skin_BoardMain + "<tr><td colspan='5'><table border='0' width='100%'>"
  Skin_BoardMain = Skin_BoardMain + "<tr height='24'>"
  Skin_BoardMain = Skin_BoardMain + "<td width='48' background='images/tile_sub.gif'>　</td>"
  Skin_BoardMain = Skin_BoardMain + "<td background='images/tile_sub.gif' align='center'><b>板块</b></td>"
  Skin_BoardMain = Skin_BoardMain + "<td width='40' background='images/tile_sub.gif' align='center'><b>主题</b></td>"
  Skin_BoardMain = Skin_BoardMain + "<td width='40' background='images/tile_sub.gif' align='center'><b>回复</b></td>"
  Skin_BoardMain = Skin_BoardMain + "<td width='240' background='images/tile_sub.gif' align='center'><b>最新发表</b></td>"
  Skin_BoardMain = Skin_BoardMain + "</tr>"
  Skin_BoardMain = Skin_BoardMain + "</table></td></tr>"
  Skin_BoardMain = Skin_BoardMain + "<tr><td colspan='5'><span id=MyBBS_Board{BoardID}></span></td></tr>"
  Skin_BoardMain = Skin_BoardMain + "<tr><td colspan='5' bgcolor='#BCD0ED'>　</td></tr>"
  Skin_BoardMain = Skin_BoardMain + "</table><br>"
  
'---------------------------首页分板块的默板：需要替换的参数为, BoardName, BoardIntro, BoardMaster, BoardTopic, BoardReply, BoardLastPostTime, BoardLastPostUser, BoardLastPostTopic
  dim Skin_BoardSecond
  Skin_BoardSecond = "<table border='0' width='100%'>"
  Skin_BoardSecond = Skin_BoardSecond + "<tr>"
  Skin_BoardSecond = Skin_BoardSecond + "<td width='48'  bgcolor='#F6F6F6'><img src='images/bf_nonew.gif' border='0'>　</td>"
  Skin_BoardSecond = Skin_BoardSecond + "<td bgcolor='#F6F6F6'>{BoardName}<br>{BoardIntro}<br>版主:{BoardMaster}</td>"
  Skin_BoardSecond = Skin_BoardSecond + "<td width='40' bgcolor='#F6F6F6' align='center'>{BoardTopic}</td>"
  Skin_BoardSecond = Skin_BoardSecond + "<td width='40' bgcolor='#F6F6F6' align='center'>{BoardReply}</td>"
  Skin_BoardSecond = Skin_BoardSecond + "<td width='240' bgcolor='#F6F6F6'>{BoardLastPostTime}<br>主题:&gt;&gt;{BoardLastPostUser}<br>发表:{BoardLastPostTopic}</td>"
  Skin_BoardSecond = Skin_BoardSecond + "</tr></table>"
  
'---------------------------帖子板块的默板：需要替换的参数为, BoardName
 dim Skin_ListMain
  Skin_ListMain = "<table border='0' width='100%'>"
  Skin_ListMain = Skin_ListMain + "<tr><td width='100%' colspan='5' background='images/tile_back.gif' height='36'>　<img src='images/nav_m.gif' border='0'>&nbsp;<font color='#FFFFFF'>{BoardName}</font></td></tr>"
  Skin_ListMain = Skin_ListMain + "<tr><td colspan='5'><table border='0' width='100%'>"
  Skin_ListMain = Skin_ListMain + "<tr height='24'>"
  Skin_ListMain = Skin_ListMain + "<td width='24' background='images/tile_sub.gif'>　</td>"
  Skin_ListMain = Skin_ListMain + "<td width='24' background='images/tile_sub.gif'>　</td>"
  Skin_ListMain = Skin_ListMain + "<td background='images/tile_sub.gif'><b>主题名称</b></td>"
  Skin_ListMain = Skin_ListMain + "<td width='80' background='images/tile_sub.gif' align='center'><b>主题发起人</b></td>"
  Skin_ListMain = Skin_ListMain + "<td width='40' background='images/tile_sub.gif' align='center'><b>回复</b></td>"
  Skin_ListMain = Skin_ListMain + "<td width='40' background='images/tile_sub.gif' align='center'><b>点击</b></td>"
  Skin_ListMain = Skin_ListMain + "<td width='200' background='images/tile_sub.gif'><b>最后发表</b></td>"
  Skin_ListMain = Skin_ListMain + "</tr>"
  Skin_ListMain = Skin_ListMain + "</table></td></tr>"
  Skin_ListMain = Skin_ListMain + "<tr><td colspan='7'><span id='MyBBS_List'></span> </td></tr>"
  Skin_ListMain = Skin_ListMain + "<tr><td colspan='7' bgcolor='#BCD0ED'>　</td></tr>"
  Skin_ListMain = Skin_ListMain + "</table><br>"  
  
'---------------------------帖子板块的默板：需要替换的参数为, BBSTitle, BBSPostUser, BBSReply, BBSView, BBSLastPostTime, BBSLastPostUser
  dim Skin_ListSecond
  Skin_ListSecond = "<table border='0' width='100%'>"
  Skin_ListSecond = Skin_ListSecond + "<tr height='24'>"
  Skin_ListSecond = Skin_ListSecond + "<td width='24' background='images/tile_sub.gif'>　</td>"
  Skin_ListSecond = Skin_ListSecond + "<td width='24' background='images/tile_sub.gif'>　</td>"
  Skin_ListSecond = Skin_ListSecond + "<td background='images/tile_sub.gif'>{BBSTitle}</td>"
  Skin_ListSecond = Skin_ListSecond + "<td width='80'  background='images/tile_sub.gif' align='center'>{BBSPostUser}</td>"
  Skin_ListSecond = Skin_ListSecond + "<td width='40'  background='images/tile_sub.gif' align='center'>{BBSReply}</td>"
  Skin_ListSecond = Skin_ListSecond + "<td width='40'  background='images/tile_sub.gif' align='center'>{BBSView}</td>"
  Skin_ListSecond = Skin_ListSecond + "<td width='200' background='images/tile_sub.gif'>{BBSLastPostTime}<br>最后发表：{BBSLastPostUser}</td>"
  Skin_ListSecond = Skin_ListSecond + "</tr></table>" 
  
  
'---------------------------管理模块首页主板块的默板：需要替换的参数为, BoardName, BoardID
 dim Skin_AdminBoardMain		
  Skin_AdminBoardMain = "<table border='0' width='100%'>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<tr><td colspan=5><table border='0' width='100%'>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<tr><td width='200' background='images/tile_back.gif' height='36'>　<img src='images/nav_m.gif' border='0'>&nbsp;<font color='#FFFFFF'>{BoardName}</font></td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td width='20'  background='images/tile_back.gif'>{BoardOrder}</td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td width='60' background='images/tile_back.gif' align='center'>{BoardSetting}</td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td background='images/tile_back.gif'>{BoardDel}</td></tr>"

  Skin_AdminBoardMain = Skin_AdminBoardMain + "<tr height='24'>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td background='images/tile_sub.gif' align='center'><b>板块名称</b></td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td background='images/tile_sub.gif' align='center'><b>顺序</b></td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td background='images/tile_sub.gif' align='center'>&nbsp;</td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td background='images/tile_sub.gif' align='center'>&nbsp;</b></td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "</tr>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "</table></td></tr>"
  
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<tr><td colspan=5><span id=MyBBS_Board{BoardID}></span></td></tr>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "</table><br>"
  
  
'---------------------------管理模块首页分板块的默板：需要替换的参数为, BoardName, BoardIntro, BoardMaster, BoardTopic, BoardReply, BoardLastPostTime, BoardLastPostUser, BoardLastPostTopic
  dim Skin_AdminBoardSecond
  Skin_AdminBoardSecond = "<table border='0' width='100%'>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<tr height='24'>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<td width='200' bgcolor='#F6F6F6'>{BoardName}</td>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<td width='20' bgcolor='#F6F6F6'>{BoardOrder}</td>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<td width='60' bgcolor='#F6F6F6' align='center'>{BoardSetting}</td>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<td bgcolor='#F6F6F6'>{BoardDel}</td>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "</tr></table>"  
%>