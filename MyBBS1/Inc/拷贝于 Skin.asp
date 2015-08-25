<%
'---------------------------首页主板块的默板：需要替换的参数为, BoardName, BoardID
 dim Skin_BoardMain		
  Skin_BoardMain = "<table border='0' width='100%'>"
  Skin_BoardMain = Skin_BoardMain + "<tr><td width='100%' colspan='5' background='images/tile_back.gif' height='36'>　<img src='images/nav_m.gif' border='0'>&nbsp;<font color='#FFFFFF'>{BoardName}</font></td></tr>"

  Skin_BoardMain = Skin_BoardMain + "<tr height='24'>"
  Skin_BoardMain = Skin_BoardMain + "<td width='48' background='images/tile_sub.gif'>　</td>"
  Skin_BoardMain = Skin_BoardMain + "<td background='images/tile_sub.gif' align='center'><b>板块</b></td>"
  Skin_BoardMain = Skin_BoardMain + "<td width='40' background='images/tile_sub.gif' align='center'><b>主题</b></td>"
  Skin_BoardMain = Skin_BoardMain + "<td width='40' background='images/tile_sub.gif' align='center'><b>回复</b></td>"
  Skin_BoardMain = Skin_BoardMain + "<td width='240' background='images/tile_sub.gif' align='center'><b>最新发表</b></td>"
  Skin_BoardMain = Skin_BoardMain + "</tr>"
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
  
  
'---------------------------首页主板块的默板：需要替换的参数为, BoardName, BoardID
 dim Skin_AdminBoardMain		
  Skin_AdminBoardMain = "<table border='0' width='100%'>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<tr><td width='25%' background='images/tile_back.gif' height='36'>　<img src='images/nav_m.gif' border='0'>&nbsp;<font color='#FFFFFF'>{BoardName}</font></td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td width='6%'  background='images/tile_back.gif'>{BoardOrder}</td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td width='8% ' background='images/tile_back.gif'>{BoardSetting}</td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td width='61%' background='images/tile_back.gif'>{BoardDel}</td></tr>"

  Skin_AdminBoardMain = Skin_AdminBoardMain + "<tr height='24'>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td width='25%' background='images/tile_sub.gif' align='center'><b>板块名称</b></td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td width='6%' background='images/tile_sub.gif' align='center'><b>顺序</b></td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td width='8%' background='images/tile_sub.gif' align='center'>&nbsp;</td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td width='61%' background='images/tile_sub.gif' align='center'>&nbsp;</b></td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "</tr>"
  
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<tr><td colspan=5><span id=MyBBS_Board{BoardID}></span></td></tr>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "</table><br>"
  
  
'---------------------------首页分板块的默板：需要替换的参数为, BoardName, BoardIntro, BoardMaster, BoardTopic, BoardReply, BoardLastPostTime, BoardLastPostUser, BoardLastPostTopic
  dim Skin_AdminBoardSecond
  Skin_AdminBoardSecond = "<table border='0' width='100%'>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<tr height='24'>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<td width='25%' bgcolor='#F6F6F6'>{BoardName}</td>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<td width='6% ' bgcolor='#F6F6F6'>{BoardOrder}</td>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<td width='8% ' bgcolor='#F6F6F6'>{BoardSetting}</td>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<td width='61%' bgcolor='#F6F6F6'>{BoardDel}</td>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "</tr></table>"  
%>