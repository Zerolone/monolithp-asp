<% option explicit%>
<!-- #include file="connection.asp" -->
<html>
<head>
<meta http-equiv="pragma" content="no-cache">
<link rel="stylesheet" type="text/css" href="main.css">
<script language="JavaScript">
	// Create an instance of our Global Variables for reference by other frames...
	var whole=new globalVars();
	function globalVars(){
		// Sets the global variables for the script. 
		// These may be changed to quickly customize the tree's apperance
		// Fonts
		// ID of selected item
		this.selid=0;
		// Other Flags
		this.homeurl=top.location.href;
	}

	function selectItem(item){
		this.selid=item;
	}

	function InitParentTree(){
		this.id=0;
		this.parentid=0;
		this.arrayid=0;
		this.title="";
		this.alt="";
       	this.href="http://www.hotcool.net/";
       	this.logo="";        
		this.son="F";
		this.open="F";
	    this.sonTree=new Array();
	}
	
	function InitSonTree(){
		this.id=0;
		this.parentid=0;
		this.arrayid=0;
		this.title="";
		this.alt="";
        this.href="http://www.hotcool.net";        
        this.logo="";        
		this.son="F";
		this.open="F";
	}
 
    //È¡µÃÄ¿Â¼Ê÷
	TreeLevel1=new Array();
<%	dim rsBoard,rsBoardType,strSQL,i,j,LastType,recNo,recNo1,myconn1
	set myconn1=server.createobject("ADODB.CONNECTION")
	myconn1.open connstr

	set rsBoard=Server.CreateObject("ADODB.Recordset")
	set rsBoardType=Server.CreateObject("ADODB.Recordset")
	rsBoardType.Open "select * from BoardType",myconn,1
	recNo=rsBoardType.RecordCount
	LastType=-1
%>
<%	for i=0 to recNo-1 %>
TreeLevel1[<%=i%>]= new InitParentTree();
TreeLevel1[<%=i%>].id=<%=rsBoardType("id")%>;
TreeLevel1[<%=i%>].parentid=0;
TreeLevel1[<%=i%>].arrayid=<%=i%>;
TreeLevel1[<%=i%>].title="<%=rsBoardType("name")%>";
TreeLevel1[<%=i%>].href="ListTopic.asp?pageNO=1&bid=<%=rsBoardType("id")%>";
TreeLevel1[<%=i%>].son="T";
TreeLevel1[<%=i%>].logo="./image/folder.gif";
<%			strSQL="select * from Board where boardtype="+cstr(rsBoardType("id"))
			rsBoard.Open strSQL, myconn1,1
		recNo1=rsBoard.RecordCount
		for j=0 to recNo1-1 %>
TreeLevel1[<%=i%>].sonTree[<%=j%>]= new InitSonTree();
TreeLevel1[<%=i%>].sonTree[<%=j%>].id=<%=rsBoard("id")%>;
TreeLevel1[<%=i%>].sonTree[<%=j%>].pid=<%=rsBoardType("id")%>;
TreeLevel1[<%=i%>].sonTree[<%=j%>].arrayid=<%=j%>;
TreeLevel1[<%=i%>].sonTree[<%=j%>].title="<%=rsBoard("cname")%>";
TreeLevel1[<%=i%>].sonTree[<%=j%>].href="ListTopic.asp?pageNO=1&bid=<%=rsBoard("id")%>";
TreeLevel1[<%=i%>].sonTree[<%=j%>].logo="./image/reddot.gif";
<%			rsBoard.moveNext
		next 
		rsBoard.Close
		rsBoardType.MoveNext
	next
	rsBoardType.Close
	session("MakeTree")=1
%>
</script>
</head>
<body class=clblack>
</body>
</html>