<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="connection.asp" -->
<!-- #include file="string.asp" -->

<%
'	POST.ASP
'	  参数：BookID  文章的序号
'          Method  方法(NEW,EDIT,REPLY),(SAVENEW,SAVEEDIT,SAVEREPLY),(DELETE)
'
'
%>

<%  dim strMethod,strTitle,strContent,rsBook,strSQL,iFid,iId,method,pmethod,bid,pgNo,rsUserInfo,strPre

	strTitle=""
	strContent=""
	iFid=0
	iId=0
	bid=request("bid")
	method=request("method")
	pmethod=request("pmethod")
	pgNo=request("PageNo")
    strMethod=UCASE(trim(request("pmethod")))
	set rsBook=SERVER.CreateObject("ADODB.RECORDSET")
	set rsUserInfo=SERVER.CreateObject("ADODB.RECORDSET")


	iId=clng(request("id"))
	if strMethod="REPLY" or strMethod="EDIT" then
	  strSQL="select * from Book where id="+request("id")
	  rsBook.Open strSQL,myconn,1
	  if rsBook.RecordCount>0 then
	     strTitle=rsBook("title")
		 strContent=rsBook("content")
 		 if strMethod="REPLY" then 
			 strPre=""
			 rsUserInfo.open "select * from UserInfo where userid='"+rsBook("userid")+"'",myconn,1
			 if rsUserInfo.RecordCount >0 then
		        strPre="【 "+rsBook("userid")+"("+rsUserInfo("nickname")+")在大作中谈到：】 "+chr(13)+chr(10)
			 end if 
			 rsUserInfo.close
		     strTitle="RE: "+strTitle
			 strContent=strPre+":"+replace(strContent,chr(10),chr(10)+":")
		 end if
		 iFid=rsBook("fid")
	  end if
	  rsBook.close
    end if	
	if strMethod="NEW" or strMethod="REPLY" or strMethod="EDIT" then
	    strMethod="SAVE"+strMethod
%>
<HTML>
<link rel="stylesheet" type="text/css" href="main.css">
<SCRIPT LANGUAGE="JavaScript">
<!--//
function check()
{
    if (document.postit.title.value.length<1)
		alert("文章标题不能为空");
	else if (document.postit.title.value.length>25) 
		alert("文章标题最长30个字");
	else if (document.postit.content.value.length<1) 
		{	alert("文章内容不能为空");document.postit.content.value="[无内容]";}
	else if (document.postit.content.value.length>16384) 
		alert("文章最长限制为16K");
	else    document.postit.submit();
}
//-->
</SCRIPT>
<BODY class=clblue>
<br><br><FORM name=postit ACTION="POST.ASP" method=post >
<table class=post><font color="#FFFFFF">文章主题：</font><font color="#FFFFFF"> 
<INPUT TYPE="text"  class=post NAME="title" size=60 value='<% =strTitle %>'>
</font>
<tr>
  <td>&nbsp;</td>
</tr>
<tr> 
  <td valign=top><br>
    <font color="#FFFFFF">文章内容:</font></td>
  <td><br>
    <TEXTAREA class=post NAME="content" rows="16" cols="60" wrap="virtual"><% =strContent %></TEXTAREA>
	<INPUT TYPE="HIDDEN" NAME="bid" value= <%=request("bid")%> >
	<INPUT TYPE="HIDDEN" NAME="ran" value= <%=request("ran")%> >
	<INPUT TYPE="HIDDEN" NAME="id" value= <%=iId %> >
	<INPUT TYPE="HIDDEN" NAME="fid" value= <%=iFid %> >
	<INPUT TYPE="HIDDEN" NAME="pageNo" value= <%=pgNo %> >
	<INPUT TYPE="HIDDEN" NAME="method" value= <%=method%> >
	<INPUT TYPE="HIDDEN" NAME="pmethod" value="<%=strMethod%>"></td>
</tr>
<tr><td></td><td>&nbsp;&nbsp;
<input type=button name="sub" value="发表文章" class=post onClick="check()">&nbsp;&nbsp;&nbsp;&nbsp;
<input type=reset name="saub" value="重新发表" class=post></td></tr>
</FORM>
</BODY></HTML>
<%  else
      if strMethod="SAVENEW" or strMethod="SAVEREPLY" or strMethod="SAVEEDIT" then
		strTitle=strFilter(request("title"))
		strContent=strFilter(request("content"))
		strContent=AddCRLF(strContent,7)
		iFid=cstr(request("fid"))
		if strMethod="SAVENEW" then
			strContent=strContent+chr(13)+chr(10)+"------------"+chr(13)+chr(10)+session("sign")
			strSQL="insert into Book(title,type,wdate,lastdate,userid,hits,content,bid,fid) values('"
			strSQL=strSQL+strTitle+"',0,now(),now(),'"
			strSQL=strSQL+session("userid")+"',0,'"
			strSQL=strSQL+strContent+"',"
			strSQL=strSQL+cstr(bid)+",0)"
			rsBook.Open strSQL,myconn
			strSQL="select * from Book where fid=0 and type=0 "
			rsBook.Open strSQL,myconn
			iFid=Cstr(rsBook("id"))
			rsBook.Close
			strSQL="Update Book set fid=id where fid=0 and type=0 "
			rsBook.Open strSQL,myconn
			strSQL="Update UserInfo set iBook=iBook+1,iPerience =iPerience+5 where userid='"+session("userid")+"'"
			rsBook.Open strSQL,myconn
			strSQL="Update Board set iBook=iBook+1 where id="+request("bid")
			rsBook.Open strSQL,myconn
		end if
		if strMethod="SAVEREPLY" then
			strContent=strContent+chr(13)+chr(10)+"------------"+chr(13)+chr(10)+session("sign")
			strSQL="insert into Book(title,type,wdate,lastdate,userid,hits,content,bid,fid) values('"
			strSQL=strSQL+strTitle+"',1,now(),now(),'"
			strSQL=strSQL+session("userid")+"',0,'"
			strSQL=strSQL+strContent+"',"
			strSQL=strSQL+cstr(bid)+","+cstr(request("fid"))+")"
			rsBook.Open strSQL,myconn
			strSQL="Update Book set lastdate=now(),reply=reply+1 where id="+request("fid")
			rsBook.Open strSQL,myconn
			strSQL="Update UserInfo set iBook=iBook+1,iPerience =iPerience+3 where userid='"+session("userid")+"'"
			rsBook.Open strSQL,myconn
			strSQL="Update Board set iBook=iBook+1 where id="+request("bid")
			rsBook.Open strSQL,myconn
		end if
		
		if strMethod="SAVEEDIT" then

			strSQL="Update Book Set lastdate=now(),title='"+strTitle+"' , content='"+strContent+"' where id=" +request("id")
			rsBook.Open strSQL,myconn
		end if
		strSQL="showdoc.asp?method="+method+"&pageNo="+pgNo+"&bid="+request("bid")+"&id="+ifid+"&ran="+request("ran")
		response.redirect strSQL
	 else
		if strMethod="DELETE" then
			if request("fid")<>request("id") then
				strSQL="delete from book where id="+request("id")
				rsBook.Open strSQL,myconn
				strSQL="update book set reply = reply-1 where id="+request("fid")
				rsBook.Open strSQL,myconn
			else
				strSQL="select * from book where fid="+request("id") +" order by type ,wdate"
				rsBook.Open strSQL,myconn,3, 3
				dim recNo,fida,hits,reply,wdate
				dim fields(2) 
				dim values(2) 
				recNo=rsBook.RecordCount
				if recNo>1 then
					rsBook.MoveFirst
					hits=rsBook("hits")
					reply=rsBook("reply")-1
					wdate=rsBook("wdate")
					rsBook.MoveNext
					fida=rsBook("id")
					rsBook("type")=0
					rsBook("fid")=rsBook("id")
					rsBook("hits")=hits
					rsbook("reply")=reply
					rsbook("wdate")=wdate
					rsBook.Update
					rsBook.close
					strSQL="Update book set fid="+cstr(fida)+" where fid="+request("id")
					rsBook.Open strSQL,myconn
				else
					rsBook.Close
				end if
				strSQL="delete from book where id="+request("id")
				rsBook.Open strSQL,myconn
				strSQL="Update Board set iBook=iBook-1 where id="+request("bid")
				rsBook.Open strSQL,myconn

			end if
		end if
		strSQL="listTopic.asp?method="+method+"&bid="+request("bid")+"&pageNo="+pgNo
		response.redirect strSQL
	  end if
	end if
%>
<%
	myconn.Close
	set rsUserInfo=nothing
	set rsBook=nothing
%>