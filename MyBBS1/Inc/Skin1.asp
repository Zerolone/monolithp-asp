<%
'---------------------------��ҳ������Ĭ�壺��Ҫ�滻�Ĳ���Ϊ, BoardName, BoardID
 dim Skin_BoardMain		
  Skin_BoardMain = "<table border='0' width='100%'>"
  Skin_BoardMain = Skin_BoardMain + "<tr><td colspan='5' background='images/tile_back.gif' height='36'>��<img src='images/nav_m.gif' border='0'>&nbsp;<font color='#FFFFFF'>{BoardName}</font></td></tr>"
  Skin_BoardMain = Skin_BoardMain + "<tr><td colspan='5'><span id=MyBBS_Board{BoardID}></span></td></tr>"
  Skin_BoardMain = Skin_BoardMain + "<tr><td colspan='5' bgcolor='#BCD0ED'>��</td></tr>"
  Skin_BoardMain = Skin_BoardMain + "</table><br>"
  
'---------------------------��ҳ�ְ���Ĭ��
  dim Skin_BoardSecond
  Skin_BoardSecond = "<table border='0' width='100%'>"
  Skin_BoardSecond = Skin_BoardSecond + "<tr height='24'>"
  Skin_BoardSecond = Skin_BoardSecond + "<td width='6%' background='images/tile_sub.gif'>��</td>"
  Skin_BoardSecond = Skin_BoardSecond + "<td width='25%' background='images/tile_sub.gif' align='center'><b>���</b></td>"
  Skin_BoardSecond = Skin_BoardSecond + "<td width='22%' background='images/tile_sub.gif' align='center'><b>����</b></td>"
  Skin_BoardSecond = Skin_BoardSecond + "<td width='22%' background='images/tile_sub.gif' align='center'><b>�ظ�</b></td>"
  Skin_BoardSecond = Skin_BoardSecond + "<td width='22%' background='images/tile_sub.gif' align='center'><b>���·���</b></td>"
  Skin_BoardSecond = Skin_BoardSecond + "</tr>"
  Skin_BoardSecond = Skin_BoardSecond + "</table>"
  
'---------------------------��ҳ�ְ��ѭ����Ĭ�壺��Ҫ�滻�Ĳ���Ϊ, BoardName, BoardIntro, BoardMaster, BoardTopic, BoardReply, BoardLastPostTime, BoardLastPostUser, BoardLastPostTopic
  dim Skin_BoardLoop	
  Skin_BoardLoop = "<table border='0' width='100%'>"
  Skin_BoardLoop = Skin_BoardLoop + "<tr height='24'>"
  Skin_BoardLoop = Skin_BoardLoop + "<td width='6%'  bgcolor='#F6F6F6'><img src='images/bf_nonew.gif' border='0'>��</td>"
  Skin_BoardLoop = Skin_BoardLoop + "<td width='25%' bgcolor='#F6F6F6'>{BoardName}<br>{BoardIntro}<br>����:{BoardMaster}</td>"
  Skin_BoardLoop = Skin_BoardLoop + "<td width='22%' bgcolor='#F6F6F6'>{BoardTopic}</td>"
  Skin_BoardLoop = Skin_BoardLoop + "<td width='22%' bgcolor='#F6F6F6'>{BoardReply}</td>"
  Skin_BoardLoop = Skin_BoardLoop + "<td width='22%' bgcolor='#F6F6F6'>{BoardLastPostTime}<br>����:&gt;&gt;{BoardLastPostUser}<br>����:{BoardLastPostTopic}</td>"
  Skin_BoardLoop = Skin_BoardLoop + "</tr></table>"
%>