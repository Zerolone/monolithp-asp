<%
'---------------------------��ҳ������Ĭ�壺��Ҫ�滻�Ĳ���Ϊ, BoardName, BoardID
 dim Skin_BoardMain		
  Skin_BoardMain = "<table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#939393'>"
  Skin_BoardMain = Skin_BoardMain + "<tr>"
  Skin_BoardMain = Skin_BoardMain + "<td bgcolor='#FFFFFF'>"
  Skin_BoardMain = Skin_BoardMain + "<table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#FFFFFF'>"
  Skin_BoardMain = Skin_BoardMain + "<tr><td height='32' colspan='5' background='{StyleBgImage}'>��<img src='images/nav_m.gif' border='0'>&nbsp;<font color='#000000'><b>{BoardName}</b></font></td></tr>"
  Skin_BoardMain = Skin_BoardMain + "</table></td></tr>"
  Skin_BoardMain = Skin_BoardMain + "<tr><td colspan='5'>"
  Skin_BoardMain = Skin_BoardMain + "<table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#FFFFFF'>"
  Skin_BoardMain = Skin_BoardMain + "<tr height='24' bgcolor='{StyleTr1}'>"
  Skin_BoardMain = Skin_BoardMain + "<td width='48'>��</td>"
  Skin_BoardMain = Skin_BoardMain + "<td align='center'><b>���</b></td>"
  Skin_BoardMain = Skin_BoardMain + "<td width='40' align='center'><b>����</b></td>"
  Skin_BoardMain = Skin_BoardMain + "<td width='40' align='center'><b>�ظ�</b></td>"
  Skin_BoardMain = Skin_BoardMain + "<td width='240' align='center'><b>���·���</b></td>"
  Skin_BoardMain = Skin_BoardMain + "</tr>"
  Skin_BoardMain = Skin_BoardMain + "</table>"
  Skin_BoardMain = Skin_BoardMain + "</td></tr>"
  Skin_BoardMain = Skin_BoardMain + "<tr><td colspan='5'><span id=MyBBS_Board{BoardID}></span></td></tr>"
  Skin_BoardMain = Skin_BoardMain + "<tr><td colspan='5' bgcolor='{StyleFoot}'>��</td></tr>"
  Skin_BoardMain = Skin_BoardMain + "</table><br>"
  
'---------------------------��ҳ�ְ���Ĭ�壺��Ҫ�滻�Ĳ���Ϊ, BoardName, BoardIntro, BoardMaster, BoardTopic, BoardReply, BoardLastPostTime, BoardLastPostUser, BoardLastPostTopic
  dim Skin_BoardSecond
  Skin_BoardSecond = "<table width='100%' border='0' cellpadding='1' cellspacing='1' name='table{TableName}'>"
  Skin_BoardSecond = Skin_BoardSecond + "<tr>"
  Skin_BoardSecond = Skin_BoardSecond + "<td width='48'><img src='images/bf_nonew.gif' border='0'>��</td>"
  Skin_BoardSecond = Skin_BoardSecond + "<td>{BoardName}<br>{BoardIntro}<br>����:{BoardMaster}</td>"
  Skin_BoardSecond = Skin_BoardSecond + "<td width='40' align='center'>{BoardTopic}</td>"
  Skin_BoardSecond = Skin_BoardSecond + "<td width='40' align='center'>{BoardReply}</td>"
  Skin_BoardSecond = Skin_BoardSecond + "<td width='240'>{BoardLastPostTime}<br>����:&gt;&gt;{BoardLastPostUser}<br>����:{BoardLastPostTopic}</td>"
  Skin_BoardSecond = Skin_BoardSecond + "</tr></table>"
  
'---------------------------���Ӱ���Ĭ�壺��Ҫ�滻�Ĳ���Ϊ, BoardName
 dim Skin_ListMain
  Skin_ListMain	   = "<table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#939393'>"
  Skin_ListMain	   = Skin_ListMain + "<tr>"
  Skin_ListMain	   = Skin_ListMain + "<td bgcolor='#FFFFFF'>"
  Skin_ListMain	   = Skin_ListMain + "<table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#FFFFFF'>"
  Skin_ListMain	   = Skin_ListMain + "<tr><td height='32' colspan='7' background='{StyleBgImage}'>��<img src='images/nav_m.gif' border='0'>&nbsp;<font color='#000000'><b>{BoardName}</b></font></td></tr>"
  Skin_ListMain	   = Skin_ListMain + "</table></td></tr>"
  Skin_ListMain	   = Skin_ListMain + "<tr><td colspan='7'>"
  Skin_ListMain	   = Skin_ListMain + " <table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#FFFFFF'>"
  Skin_ListMain	   = Skin_ListMain + "<tr height='24' bgcolor='{StyleTr1}'>"
  Skin_ListMain	   = Skin_ListMain + "<td width='24'>��</td>"
  Skin_ListMain	   = Skin_ListMain + "<td width='24'>��</td>"
  Skin_ListMain	   = Skin_ListMain + "<td><b>��������</b></td>"
  Skin_ListMain	   = Skin_ListMain + "<td width='80' align='center'><b>���ⷢ����</b></td>"
  Skin_ListMain	   = Skin_ListMain + "<td width='40' align='center'><b>�ظ�</b></td>"
  Skin_ListMain	   = Skin_ListMain + "<td width='40' align='center'><b>���</b></td>"
  Skin_ListMain	   = Skin_ListMain + "<td width='200'><b>��󷢱�</b></td>"
  Skin_ListMain	   = Skin_ListMain + "</tr>"
  Skin_ListMain	   = Skin_ListMain + "</table>"
  Skin_ListMain	   = Skin_ListMain + "</td></tr>"
  Skin_ListMain	   = Skin_ListMain + "<tr><td colspan='7'><span id='MyBBS_List'></span></td></tr>"
  Skin_ListMain	   = Skin_ListMain + "<tr><td colspan='7' bgcolor='{StyleFoot}'>��</td></tr>"
  Skin_ListMain	   = Skin_ListMain + "</table><br>  "


'---------------------------���Ӱ���Ĭ�壺��Ҫ�滻�Ĳ���Ϊ, BBSTitle, BBSPostUser, BBSReply, BBSView, BBSLastPostTime, BBSLastPostUser
  dim Skin_ListSecond
  Skin_ListSecond = "<table width='100%' border='0' cellpadding='1' cellspacing='1' name='tablelist'>"
  Skin_ListSecond = Skin_ListSecond + "<tr height='24'>"
  Skin_ListSecond = Skin_ListSecond + "<td width='24'>��</td>"
  Skin_ListSecond = Skin_ListSecond + "<td width='24'>��</td>"
  Skin_ListSecond = Skin_ListSecond + "<td >{BBSTitle}</td>"
  Skin_ListSecond = Skin_ListSecond + "<td width='80'   align='center'>{BBSPostUser}</td>"
  Skin_ListSecond = Skin_ListSecond + "<td width='40'   align='center'>{BBSReply}</td>"
  Skin_ListSecond = Skin_ListSecond + "<td width='40'   align='center'>{BBSView}</td>"
  Skin_ListSecond = Skin_ListSecond + "<td width='200'>{BBSLastPostTime}<br>��󷢱�{BBSLastPostUser}</td>"
  Skin_ListSecond = Skin_ListSecond + "</tr>"
  Skin_ListSecond = Skin_ListSecond + "</table>"  
  

'---------------------------���Ӱ���Ĭ�壺��Ҫ�滻�Ĳ���Ϊ, BBSTitle, UserName, Topic, BBSBody, BBSReply, BBSView, BBSPostTime
 dim Skin_ListViewMain
  Skin_ListViewMain	   = "<table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#939393'>"
  Skin_ListViewMain	   = Skin_ListViewMain + "<tr>"
  Skin_ListViewMain	   = Skin_ListViewMain + "<td bgcolor='#FFFFFF'>"
  Skin_ListViewMain	   = Skin_ListViewMain + "<table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#FFFFFF'>"
  Skin_ListViewMain	   = Skin_ListViewMain + "<tr><td height='32' colspan='2' background='{StyleBgImage}'>��<img src='images/nav_m.gif' border='0'>&nbsp;<font color='#000000'><b>{BBSTitle}</b></font></td></tr>"
  Skin_ListViewMain	   = Skin_ListViewMain + "</table></td></tr>"
  Skin_ListViewMain	   = Skin_ListViewMain + "<tr><td colspan='2'>"
  Skin_ListViewMain	   = Skin_ListViewMain + " <table width='100%' border='0' cellpadding='0' cellspacing='1' bgcolor='#FFFFFF'>"
  Skin_ListViewMain	   = Skin_ListViewMain + "<tr height='24' bgcolor='{StyleTr1}'>"
  Skin_ListViewMain	   = Skin_ListViewMain + "<td width='150'>���ID��{ID}<br>�û�����{UserName}<br>��������{Topic}</td>"
  Skin_ListViewMain	   = Skin_ListViewMain + "<td>{BBSBody}<br>--------------------<br>�ظ�����{BBSReply}  ���������{BBSView}  ����ʱ�䣺{BBSPostTime}</td>"
  Skin_ListViewMain	   = Skin_ListViewMain + "</tr>"
  Skin_ListViewMain	   = Skin_ListViewMain + "</table>"
  Skin_ListViewMain	   = Skin_ListViewMain + "</td></tr>"
  Skin_ListViewMain	   = Skin_ListViewMain + "<tr><td colspan='2'><span id='MyBBS_List'></span></td></tr>"
  Skin_ListViewMain	   = Skin_ListViewMain + "<tr><td colspan='2' bgcolor='{StyleFoot}'>��</td></tr>"
  Skin_ListViewMain	   = Skin_ListViewMain + "</table><br>  "

'---------------------------���Ӱ���Ĭ�壺��Ҫ�滻�Ĳ���Ϊ, BBSTitle, UserName, Topic, BBSBody, BBSPostTime
  dim Skin_ListViewSecond
  Skin_ListViewSecond = "<table width='100%' border='0' cellpadding='1' cellspacing='1' name='tablelist'>"
  Skin_ListViewSecond = Skin_ListViewSecond + "<tr height='24'>"
  Skin_ListViewSecond = Skin_ListViewSecond + "<td width='150'>���ID��{ID}<br>�û�����{UserName}<br>��������{Topic}</td>"
  Skin_ListViewSecond = Skin_ListViewSecond + "<td><br><b>{BBSTitle}</b><br>{BBSBody}<br>--------------------<br>����ʱ�䣺{BBSPostTime}<br></td>"
  Skin_ListViewSecond = Skin_ListViewSecond + "</tr>"
  Skin_ListViewSecond = Skin_ListViewSecond + "</table>"    
  
  
'---------------------------��̨����  
'---------------------------����ģ����ҳ������Ĭ�壺��Ҫ�滻�Ĳ���Ϊ, BoardName, BoardID
 dim Skin_AdminBoardMain		
  Skin_AdminBoardMain = "<table border='0' width='100%'>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<tr><td colspan=5><table border='0' width='100%'>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<tr><td width='200' background='images/tile_back.gif' height='36'>��<img src='images/nav_m.gif' border='0'>&nbsp;<font color='#FFFFFF'>{BoardName}</font></td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td width='20'  background='images/tile_back.gif'>{BoardOrder}</td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td width='60' background='images/tile_back.gif' align='center'>{BoardSetting}</td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td background='images/tile_back.gif'>{BoardDel}</td></tr>"

  Skin_AdminBoardMain = Skin_AdminBoardMain + "<tr height='24'>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td background='images/tile_sub.gif' align='center'><b>�������</b></td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td background='images/tile_sub.gif' align='center'><b>˳��</b></td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td background='images/tile_sub.gif' align='center'>&nbsp;</td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<td background='images/tile_sub.gif' align='center'>&nbsp;</b></td>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "</tr>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "</table></td></tr>"
  
  Skin_AdminBoardMain = Skin_AdminBoardMain + "<tr><td colspan=5><span id=MyBBS_Board{BoardID}></span></td></tr>"
  Skin_AdminBoardMain = Skin_AdminBoardMain + "</table><br>"
  
  
'---------------------------����ģ����ҳ�ְ���Ĭ�壺��Ҫ�滻�Ĳ���Ϊ, BoardName, BoardIntro, BoardMaster, BoardTopic, BoardReply, BoardLastPostTime, BoardLastPostUser, BoardLastPostTopic
  dim Skin_AdminBoardSecond
  Skin_AdminBoardSecond = "<table border='0' width='100%'>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<tr height='24'>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<td width='200' bgcolor='#F6F6F6'>{BoardName}</td>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<td width='20' bgcolor='#F6F6F6'>{BoardOrder}</td>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<td width='60' bgcolor='#F6F6F6' align='center'>{BoardSetting}</td>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "<td bgcolor='#F6F6F6'>{BoardDel}</td>"
  Skin_AdminBoardSecond = Skin_AdminBoardSecond + "</tr></table>"  
%>