<html>
<head>
<title>发送短消息</title>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<style type="text/css">
td{font-size:9pt;}
A:link {COLOR: #000000; TEXT-DECORATION: none; font-size: 9pt}
A:active {COLOR: #000000; TEXT-DECORATION: none; font-size: 9pt}
A:visited {COLOR: #000000; TEXT-DECORATION: none; font-size: 9pt}
.song9 {  font-family: 宋体; font-size: 9pt}
</style>

</head>
<script language="javascript">
<!--
   
   function saysfocus( )
    {
	      //  top.main1.scroll(0, 65000);
        //    setTimeout('scrollWindow()', 200);
        
        
        
            document.inputform.message.focus();
		
	//parent.userlist.location.reload();	
	}

-->
</script>

<body aLink="#000000" bgcolor=#99aacc link="#000000" text="#000000" topMargin="4" Link="#000000" MARGINWIDTH="5" MARGINHEIGHT="4">
<%
if request("message")<>"" then '把request("submit")改成了request("message") /by c.t. 14:46 19991020
    if trim(request("message"))<>"" then
       from=request("myself")
       
       towhom=request("to")
       message=request("message")
       
       

       mymessage=from&"&&"&towhom&"&&"&message
       Application.Lock
       application("message")=Application("message")&mymessage&"||"
       Application.UnLock
       Response.write "<p align=center>"&from&",你的信息发送成功</p>"
       %>
      <script language="javascript">
      <!--
      window.self.close();
      -->
      </script> 
    <%else
       Response.Write  "<p align=center><font color=red>发信内容不能为空</font>"
     end if
    

else
   

%>

<p align=center><b><font size="2">发信息给:</font></b><%=request("user")%></p>
<form action=sendmess.asp method=post name=inputform >
<input type=hidden name=to value=<%=request("user")%>>
<input type=hidden name=myself value=<%=request("myself")%>>
  <p align=center><b><font size="2" color="#3399CC">信息</font></b> 
    <input type=text name=message size=40 maxlength=100></p>
<p align=center><input type="submit" name=submit1 value="发送"></p>
</form>

<div align="center">
  <script language="javascript">
      <!--
      saysfocus();
      -->
      </script>
  <font size="2"><b>旅者制作</b></font> </div>
</body>
</html>
<%end if%>