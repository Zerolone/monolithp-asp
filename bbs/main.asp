<%@ LANGUAGE="VBSCRIPT" %>
<% option explicit%>
<!-- #include file="std.asp" -->
<%	OnlyLogin %>
<html>
<head>
<meta http-equiv="pragma" content="no-cache">
<title><%=Application("Title")%></title>
</head>
<script language="javascript">
	<%if( session("Msg")=1 ) then %>
		alert("<% =session("StrMsg")%>");
	<%	session("StrMsg")=" "  
		session("Msg")=0
	  end if
	%>
</script>
    
<frameset cols="128,*" framespacing="0" border="0" frameborder="0" >
		
  <frameset rows="1,60,1*" framespacing="0" border="0" frameborder="0" name="frmBigLeft"> 
    <frame name="frmTree"  scrolling="no" MARGINWIDTH="0" MARGINHEIGHT="0" src="GlobalTree.asp">
		  <frame name="frmTop"  scrolling="no" MARGINWIDTH="0" MARGINHEIGHT="0" src="online.asp">

		  <frame name="frmLeft" scrolling="auto" MARGINWIDTH="0" MARGINHEIGHT="0" src="menu_tree.asp">
		</frameset>
    	
  <frameset rows="60,1*,15" framespacing="0" border="0" frameborder="0" name="frmBigRight"> 
    <frameset cols="10,*,20" framespacing="0" border="0" frameborder="0" name="frmBigLeft">
		    <frame name="frmTitle1"  scrolling="no" MARGINWIDTH="0" MARGINHEIGHT="0" src="title1.htm">
		    <frame name="frmTitle2"  scrolling="no" MARGINWIDTH="0" MARGINHEIGHT="0" src="title2.htm">
		    <frame name="frmTitle3"  scrolling="no" MARGINWIDTH="0" MARGINHEIGHT="0" src="title3.htm">
		  </frameset> 
	      <frameset cols="2,*,7" framespacing="0" border="0" frameborder="0" name="frmBigLeft">
		    <frame name="frmTitle1"  scrolling="no" MARGINWIDTH="0" MARGINHEIGHT="0" src="main1.htm">
  		    <frame name="frmMain" src="listtopic.asp?method=new&amp;pageNo=1" noresize scrolling="auto" MARGINWIDTH="1" MARGINHEIGHT="0"> 
		    <frame name="frmTitle3"  scrolling="no" MARGINWIDTH="0" MARGINHEIGHT="0" src="main3.htm">
		  </frameset> 
	      <frameset cols="10,*,20" framespacing="0" border="0" frameborder="0" name="frmBigLeft">
		    <frame name="frmTitle1"  scrolling="no" MARGINWIDTH="0" MARGINHEIGHT="0" src="buttom1.htm">
		    <frame name="frmTitle2"  scrolling="no" MARGINWIDTH="0" MARGINHEIGHT="0" src="buttom2.htm">
		    <frame name="frmTitle3"  scrolling="no" MARGINWIDTH="0" MARGINHEIGHT="0" src="buttom3.htm">
		  </frameset>   	      
		</frameset>
		<noframes>
		<body>
		<p>This page uses frames, but your browser doesn't support them.</p>
		</body>
		</noframes>
</frameset>
</html>
