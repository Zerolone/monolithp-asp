<html>
<head>
<title>���Ͷ���Ϣ</title>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<style type="text/css">
td{font-size:9pt;}
A:link {COLOR: #000000; TEXT-DECORATION: none; font-size: 9pt}
A:active {COLOR: #000000; TEXT-DECORATION: none; font-size: 9pt}
A:visited {COLOR: #000000; TEXT-DECORATION: none; font-size: 9pt}
.song9 {  font-family: ����; font-size: 9pt}
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
if request("message")<>"" then '��request("submit")�ĳ���request("message") /by c.t. 14:46 19991020
    if trim(request("message"))<>"" then
       from=request("myself")
       
       towhom=request("to")
       message=request("message")
       
       

       mymessage=from&"&&"&towhom&"&&"&message
       Application.Lock
       application("message")=Application("message")&mymessage&"||"
       Application.UnLock
       Response.write "<p align=center>"&from&",�����Ϣ���ͳɹ�</p>"
       %>
      <script language="javascript">
      <!--
      window.self.close();
      -->
      </script> 
    <%else
       Response.Write  "<p align=center><font color=red>�������ݲ���Ϊ��</font>"
     end if
    

else
   

%>

<p align=center><b><font size="2">����Ϣ��:</font></b><%=request("user")%></p>
<form action=sendmess.asp method=post name=inputform >
<input type=hidden name=to value=<%=request("user")%>>
<input type=hidden name=myself value=<%=request("myself")%>>
  <p align=center><b><font size="2" color="#3399CC">��Ϣ</font></b> 
    <input type=text name=message size=40 maxlength=100></p>
<p align=center><input type="submit" name=submit1 value="����"></p>
</form>

<div align="center">
  <script language="javascript">
      <!--
      saysfocus();
      -->
      </script>
  <font size="2"><b>��������</b></font> </div>
</body>
</html>
<%end if%>