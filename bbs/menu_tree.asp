<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="std.asp" -->
<%	OnlyLogin  %>
<html>
<head>
<link rel="stylesheet" type="text/css" href="main.css">
<% if session("MakeTree")=0 then %> 
<meta http-equiv="refresh" content="10">
<% end if %> 
<style type="text/css">
<!--
.white2 {  font-size: 10pt; color: #FFFFFF; text-decoration: none}
-->
</style>
</head>
<body topmargin="0" class=clblack bgcolor="#3399CC" >
<% if session("MakeTree")=1 then %> 
<p class=menu> <font color="#FFFFFF"> 
  <script LANGUAGE="JavaScript">
    //���游�����������Ϣ
	var theParentTree=top.frmTree.TreeLevel1;
	if (theParentTree==null){document.location="menu_tree.asp";}
	if (theParentTree!=null){
	var ParentTreeLength=theParentTree.length;
	//����ȫ����Ϣ	
	var gVars=top.frmTree.whole;

	//��ʾ��Ŀ¼��
    function DispSonTree(item,lastnode){
             for (j=0; j < theParentTree[item].sonTree.length; j++) {
                    writeStr = "";
					if (j==theParentTree[item].sonTree.length-1)
							writeStr += "<IMG SRC='./image/line_06.gif'  align=absmiddle>";
                    else	writeStr += "<IMG SRC='./image/line_05.gif'  align=absmiddle>";
					//��ʾ�����LOGOͼƬ
					writeStr += "<IMG SRC='" + theParentTree[item].sonTree[j].logo + "' border=0 align=absmiddle'><A class=menu  HREF='" + theParentTree[item].sonTree[j].href + "' target='frmMain'>" + theParentTree[item].sonTree[j].title + "</A><br>";
		            document.write(writeStr);
            }   
    }
    
    //Ŀ¼����չ
    function expand(ID){
        if (top.frmTree.TreeLevel1[ID].open=="T")  
				top.frmTree.TreeLevel1[ID].open="F";
        else	top.frmTree.TreeLevel1[ID].open="T";
        tempPos=self.location.href.indexOf("?");
        if (tempPos!=-1)     Href=self.location.href.substring(0,tempPos);
        else			     Href=self.location.href;
        self.location.href=Href+"?ID="+ID;
    }

	//��ʾĿ¼��    
	for (i=0; i < ParentTreeLength; i++) {
		writeStr = "";
        dispSon=false;
        lastnode=false;
        
        //��ʾĿ¼���Ľڵ�ͼƬ		
		if (theParentTree[i].son=="T"){
//		   if (i==ParentTreeLength-1)
		   {
		      if (theParentTree[i].open=="F"){
		          writeStr += "<A class=menu href='javascript:expand(" + theParentTree[i].arrayid + ");'><IMG SRC='./image/t_plus.gif' align=absmiddle border=0></a>";
		      }
		      else{
		          writeStr += "<A class=menu  href='javascript:expand(" + theParentTree[i].arrayid + ");'><IMG SRC='./image/t_sub.gif' align=absmiddle border=0></a>";
		          dispSon=true;
		      }
		      lastnode=true //���︸Ŀ¼��ĩ�˱��    
		   }
/*		   else{
		      if (theParentTree[i].open=="F"){
		          writeStr += "<A class=menu href='javascript:expand(" + theParentTree[i].arrayid + ");'><IMG SRC='./image/t_plus.gif' align=absmiddle border=0></a>";
		      }
		      else{
		          writeStr += "<A class=menu href='javascript:expand(" + theParentTree[i].arrayid + ");'><IMG SRC='./image/t_sub.gif' align=absmiddle border=0></a>";
		          dispSon=true;
		      }    
		   }
*/		          
	    }
/*	    else{
		   if (i==ParentTreeLength-1)
					writeStr += "<IMG SRC='"+gVars.emptyImgLast+"' align=absmiddle >";
		   else		writeStr += "<IMG SRC='"+gVars.emptyImg+"' align=absmiddle >";
	    }
*/	    
	    //��ʾ�����LOGOͼƬ
	    writeStr += "<A class=menu1 href='javascript:expand(" + theParentTree[i].arrayid + ");'><IMG NAME='JP" + theParentTree[i].arrayid + "' SRC='"+theParentTree[i].logo+"' border=0 align=absmiddle>";
	    writeStr += theParentTree[i].title;
	    writeStr += "</A><br>";
	    document.write(writeStr);
        //��ʾ����Ŀ¼��
        if (dispSon){
           DispSonTree(theParentTree[i].arrayid,lastnode);
        }   
    }
}
</script>
  <IMG SRC='./image/t_ver1.gif' border=0 align=absmiddle><IMG SRC='./image/t_hr1.gif' border=0 align=absmiddle><br>
  <IMG SRC='./image/t_sub.gif' align=absmiddle border=0><A class=menu2  HREF='listtopic.asp?method=new&amp;pageNo=1' target=frmMain><span align=absmiddle><IMG SRC='./image/i_option.gif' border=0 align=absmiddle><span class="white2">��������</span></span></A><br>
  <% if lcase(session("userid"))<>"guest" then%> <IMG SRC='./image/t_sub.gif' align=absmiddle border=0><A class=menu2  HREF='adduser.asp?method=modify&userid=<%=session("userid")%>' target=frmMain><span align=absmiddle><IMG SRC='./image/i_option.gif' border=0 align=absmiddle><span class="white2">�ҵ���Ϣ</span></span></A><br>
  <% else %> <IMG SRC='./image/t_sub.gif' align=absmiddle border=0><A class=menu2  HREF='adduser.asp?method=inside' target=frmMain><span align=absmiddle><IMG SRC='./image/i_option.gif' border=0 align=absmiddle><span class="white2">��Ҫע��</span></span></A><br>
  <% end if %> <IMG SRC='./image/t_sub.gif' align=absmiddle border=0><A class=menu2  HREF='sysInfo.asp' target=frmMain><span align=absmiddle><IMG SRC='./image/i_option.gif' border=0 align=absmiddle><span class="white2">��̳��Ϣ</span></span></A><br>
  <IMG SRC='./image/t_sub.gif' align=absmiddle border=0><A class=menu2  HREF='useronline.asp' target=frmMain><span align=absmiddle><IMG SRC='./image/i_option.gif' border=0 align=absmiddle><span class="white2">����Ѱ��</span></span></A><br>
  <IMG SRC='./image/t_sub.gif' align=absmiddle border=0><A class=menu2  HREF='QueryUser.asp' target=frmMain><span align=absmiddle><IMG SRC='./image/i_option.gif' border=0 align=absmiddle><span class="white2">�����û�</span></span></A><br>
  <% if isSysOP(session("userid")) then%> <IMG SRC='./image/t_sub.gif' align=absmiddle border=0><A class=menu2  HREF='ClubManager.asp' target=frmMain><span align=absmiddle><IMG SRC='./image/i_option.gif' border=0 align=absmiddle><span class="white2">վ�����</span></span></A><br>
  <% end if %> <IMG SRC='./image/t_subl.gif' align=absmiddle border=0><A class=menu1  HREF='quit.asp' target='_top' ><span align=absmiddle><IMG SRC='./image/i_exit.gif' border=0 align=absmiddle><span class="white2">�˳���̳</span></span></A></font> 
</p>
<p class=menu align="center"><a href="mailto:info@netshi.net" class="white2">��������v1.0</a><a href="http://www.netshi.net/campus/chat/jjuyou/login.asp" target="_blank" class="white2"></a> 
  <br>
  <% end if%> <center><font color="#ffee00"><br>
</p>

<p>

 </body>
</html>
