<%@LANGUAGE="JavaScript" CODEPAGE="936"%>
<!-- #include virtual ="/Manage/Check.asp"-->
<!--
���� COCOON ASP ̽�� ����������������������������������������������
��                                                               ��
��  ��л��ʹ�� COCOON ASP ̽�� v2.x                              ��
��  ��������ȫ��������ѣ���������⸴�ơ��������޸ĺ�ʹ�ã�     ��
��  �����ù���������� ����������ҵ��;������������ʹ�����շѡ�  ��
��                                                               ��
��  ʹ��ʱ���뱣���˶���Ϣ��лл���                             ��
��                                                               ��
��  ���ߣ�Sunrise_Chen @ Cocoon sTudio.                          ��
��                                       2002/12/12              ��
��                                                               ��
����������������������������������������������  ccopus.com ��������
-->
<script language="JavaScript" runat="server">function getEngVerJs(){return ScriptEngineMajorVersion()+"."+ScriptEngineMinorVersion()+"."+ ScriptEngineBuildVersion()+" ";}</script>
<script language="VBScript" runat="server">Function getEngVerVBS() : getEngVerVBS=ScriptEngineMajorVersion() &"."&ScriptEngineMinorVersion() &"." & ScriptEngineBuildVersion() & " " : End Function</script>
<%Response.Expires=0;Response.Buffer =true;var tPageStartTime=new Date();var sObjName=Request.Form("sObjName");var sServerName=Request.ServerVariables("SERVER_NAME")(1);var sServerAddr=Request.ServerVariables("LOCAL_ADDR")(1);var sTheFile=Request.ServerVariables("URL")(1);var iServerTimeout=Server.ScriptTimeout;var sNumberOfCpu=Request.ServerVariables("NUMBER_OF_PROCESSORS")+'';var sQueryString=Request.ServerVariables("QUERY_STRING")+'';var sOsInfo=Request.ServerVariables("OS")+'';if(sOsInfo=='undefined') sOsInfo=' [δ֪]';var arrObj=new Array(Array( "MSWC.AdRotator","", 1 ),Array( "MSWC.BrowserType","", 1 ),Array( "MSWC.NextLink","", 1 ),Array( "MSWC.Tools","", 1 ),Array( "MSWC.Status","", 1 ),Array( "MSWC.Counters","", 1 ),Array( "IISSample.ContentRotator","", 1 ),Array( "IISSample.PageCounter","", 1 ),Array( "MSWC.PermissionChecker","", 1 ),Array( "WScript.Shell","", 1 ),Array( "Scripting.FileSystemObject","", 1 ),Array( "ADODB.Connection","ActiveX Data Object  [ADO]", 1 ),Array( "CDONTS.NewMail","Collaboration Data Object  [CDO]", 1 ),Array( "SoftArtisans.FileUp","SA-FileUp �ļ��ϴ�", 2 ),Array( "SoftArtisans.FileManager","SA-FM �ļ�����", 2 ),Array( "LyfUpload.UploadFile","LyfUpload �ļ��ϴ�", 2 ),Array( "Persits.Upload.1","ASPUpload �ļ��ϴ�", 2 ),Array( "w3.upload","w3 upload �ļ��ϴ�", 2 ),Array( "iismail.iismail.1","IISemail", 3 ),Array( "JMail.SMTPMail","w3 Jmail", 3 ),Array( "Persits.MailSender","ASPemail", 3 ),Array( "SMTPsvg.Mailer","ASPmail", 3 ),Array( "dkQmail.Qmail","dkQmail", 3 ),Array( "SmtpMail.SmtpMail.1","SMTPmail", 3 ),Array( "Geocel.Mailer","Geocel", 3),Array( "SoftArtisans.ImageGen","SA ��ͼ���д���", 4),Array( "W3Image.Image","Dimac ��ͼ���д���", 4));if(sObjName.Count>0) arrObj [arrObj.length]=Array( sObjName(1),"", 99 );function getObjVer(objName){try{ var objTest=Server.CreateObject(objName);}catch(e){ if(e.number==-2147221005) return "N/A";}try{ var sObjVer=objTest.Version;}catch(e){ return "";}if(isNaN(parseInt(sObjVer))) return "";objTest=null;return sObjVer;}%><html><head><title>COCOON ASP ̽��
</title><meta http-equiv="Content-Type" content="text/html;charset=gb2312"><meta http-equiv="Expires" content="0"><meta http-equiv="Cache-Control" content="no-cache"><meta http-equiv="Pragma" content="no-cache"><style type="text/css">body,td {	FONT-SIZE:9pt;FONT-FAMILY:"Tahoma","Arial","Helvetica","sans-serif" }a { COLOR:#000000;TEXT-DECORATION:none }a:hover { COLOR:#ff0000;TEXT-DECORATION:none }a.td1o2{ border:3px #333 double;padding-left:5px;padding-right:5px;}a.td2o2{ border:3px #333 double;padding-left:5px;padding-right:5px;}.tbl1 {	BORDER-RIGHT:#3f5294 1px solid;BORDER-TOP:#3f5294 1px solid;FONT-SIZE:9pt;BORDER-LEFT:#3f5294 1px solid;BORDER-BOTTOM:#3f5294 1px solid }.td1 { BORDER-RIGHT:#ffffff 0px solid;BORDER-TOP:#ffffff 1px solid;BORDER-LEFT:#ffffff 1px solid;color:#336699;BORDER-BOTTOM:#ffffff 0px solid;BACKGROUND-COLOR:#abb6dc}.tbl1o1 { BACKGROUND-COLOR:#8595cb }.td1o1 { BORDER-RIGHT:#ffffff 0px solid;BORDER-TOP:#ffffff 1px solid;BORDER-LEFT:#ffffff 1px solid;BORDER-BOTTOM:#ffffff 0px solid;BACKGROUND-COLOR:#e2e7f3}.tr1 { BACKGROUND-color:#5c72ba }.td1o2 { BACKGROUND-COLOR:#f3f4fa }.td102a { COLOR :#5C72BA }.tbl2 { BORDER:#50A0A0 1px solid;FONT-SIZE:9pt;}.td2 { BORDER-RIGHT:#50A0A0 0px solid;BORDER-TOP:#ffffff 1px solid;BORDER-LEFT:#ffffff 1px solid;COLOR:#308080;BORDER-BOTTOM:#ffffff 0px solid;BACKGROUND-COLOR:#ADCDC2 }.tbl2o1 { BACKGROUND-COLOR:#50A0A0 }.td2o1 { BORDER-RIGHT:#ffffff 0px solid;BORDER-TOP:#ffffff 1px solid;BORDER-LEFT:#ffffff 1px solid;BORDER-BOTTOM:#ffffff 0px solid;BACKGROUND-COLOR:#CDEDE2;}.tr2 { BACKGROUND-COLOR:#50A0A0 }.td2o2 { BACKGROUND-COLOR:#DDFDF2 }.PicBar { background-color:#336699;border:1px solid #000000;height:12px;}.td1o21 {BACKGROUND-COLOR:#f3f4fa;padding:7px;}</style><script language="JavaScript">
window.onerror=new Function("return true;");</script><script language="JavaScript">
function getEngVerJs(){try{return ScriptEngineMajorVersion()+"."+ScriptEngineMinorVersion()+"."+ ScriptEngineBuildVersion()+" ";}catch(e){return 'unknow';}}</script><script language="VBScript">
Function getEngVerVBS() : On Error Resume Next : getEngVerVBS=ScriptEngineMajorVersion() &"."&ScriptEngineMinorVersion() &"." & ScriptEngineBuildVersion() & " " : End Function</script><script language="JavaScript">
window.onload=new Function("document.readyState='complete';");function showLoadingAnimation(sDivName,n){var a,o;var a=Array('-','\\','|','/');if(!(o=document.getElementById(sDivName))) return ;var i=(isNaN(n)?0:n%4);o.innerHTML=a [i];if(document.readyState=='complete') return ;setTimeout('showLoadingAnimation("'+sDivName+'",'+(i+1)+')',1000);}</script><script language="JavaScript">
function getObjByID(n) {var p,i,x;var d=document;if(!(x=d [n])&&d.all) x=d.all [n];for(i=0;!x&&i<d.forms.length;i++) x=d.forms [i][n];for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=getObjByID(n,d.layers [i].document);if(!x && d.getElementById) x=d.getElementById(n);return x;}function getBrower(){var re=/(Netscape|Opera|NetCaptor|MSN |MSIE|MyIE|Mozilla\/)( [^;|\(|\)|$]+)/gim;var arr, arrResult;while((arr=re.exec(arguments [0]))!=null){arrResult=arr};if(arrResult) return arrResult [0];return navigator.appName;}function getOS(){var re=/(Windows|Mac|unix|Linux|SunOS|BSD)( [^;|\(|\)]+)/gim;var arr, arrResult;while((arr=re.exec(arguments [0]))!=null){arrResult=arr};if(arrResult) return arrResult [0];return navigator.platform;}function getMSN(){var sReturnVal,sCurrStat,iInStr;if(arguments.length<1) return "Unknow";sCurrStat=arguments [0];iInStr=sCurrStat.indexOf("MSN ");if(iInStr>0) sReturnVal=sCurrStat.substr(iInStr,sCurrStat.indexOf(";",iInStr)-iInStr);else sReturnVal="no MSN";return sReturnVal;}function getDotNet(){var sReturnVal,sCurrStat,iInStr,iInStr1;if(arguments.length<1) return "Unknow";sCurrStat=arguments [0];iInStr=sCurrStat.indexOf(".NET ");iInStr1=sCurrStat.indexOf(";",iInStr+1);if(iInStr1<0) iInStr1=sCurrStat.indexOf(")",iInStr);if(iInStr>0) sReturnVal=sCurrStat.substr(iInStr,iInStr1-iInStr);else sReturnVal="no .NET CLR";return sReturnVal;}function plugNS(sObjName){var np=navigator.plugins;for(var i=0;i<np.length;i++)if(np [i].name.indexOf(sObjName)>-1)return true;return false;}function plugIE(sObjID){if(!document.body) document.write('<body>');document.body.addBehavior("#default#clientcaps");if(document.body.isComponentInstalled("{"+sObjID+"}","componentid"))return true;return false;}function plugVerIE(sObjID){try{return document.body.getComponentVersion("{"+sObjID+"}","componentid");}catch(e){return "unknow";}}function checkPlug(plugName,plugID,n){var aClientComponent=new Array();if(isNescape){aClientComponent [0]=Array(plugName,"",(plugNS(plugName)?"��":"��"),2);}else{if(plugID) aClientComponent [0]=Array(plugName,plugID,(plugIE(plugID)?"�� ( "+plugVerIE(plugID)+" )":"��"),2);}addRow("tblClientComponent",aClientComponent,2);}function addRow(tblName,arrName,n){var newRow,newCell;var oTbl=getObjByID(tblName);if(!oTbl) return false;if(!arrName||arrName.length<1) return false;for(var i=0;i<arrName.length;i++){newRow=oTbl.insertRow(oTbl.rows.length>1?oTbl.rows.length-1:1);if(arrName [i].length<2) continue;newCell=newRow.insertCell(0);newCell.className="td"+n+"o1";newCell.innerHTML="&nbsp;"+arrName [i][0]+(arrName [i][1]?" <font color='#666666' style='font-size=9px'> [ "+arrName [i][1]+" ]</font>":"");newCell=newRow.insertCell(1);newCell.className="td"+n+"o2";newCell.align="center";newCell.innerHTML="<div style='width:200px;text-align:left'>"+arrName [i][2]+"</div>";}}var aPlug=new Array();var bJavaScript=bVBScript=false;var userAgent=navigator.userAgent;var aClientInfo=new Array();var aServerEnv=new Array();var oDiv;var isNescape=(navigator.appName.indexOf("Netscape")>-1?true:false);function pageInit(){aPlug=Array(Array("Internet Explorer","89820200-ECBD-11CF-8B85-00AA005B4383"),Array("Outlook Express","44BBA840-CC51-11CF-AAFA-00AA00B6015C"),Array("Offline Browsing Pack","3AF36230-A269-11D1-B5BF-0000F8051515"),Array("Microsoft virtual machine","08B0E5C0-4FCB-11CF-AAA5-00401C608500"),Array("Shockwave Flash","D27CDB6E-AE6D-11CF-96B8-444553540000"),Array("Shockwave for Director","2A202491-F00D-11CF-87CC-0020AFEECF20"),Array("RealPlayer","23064720-C4F8-11D1-994D-00C04F98BBC9"),Array("Windows Media Player","22D6F312-B0F6-11D0-94AB-0080C74C7E95"),Array("QuickTime",""),Array("VivoActive",""),Array("LiveAudio",""),Array("Vector Graphics Rendering(VML)","10072CEC-8CC1-11D1-986E-00A0C955B42F"),Array("DirectAnimation","283807B5-2C60-11D0-A31D-00AA00B92C03"),Array("DirectAnimation Java Classes","4F216970-C90C-11D1-B5C7-0000F8051515"),Array("Dynamic HTML Data Binding","9381D8F2-0288-11D0-9501-00AA00B911A5"),Array("Dynamic HTML Data Binding for Java","4F216970-C90C-11D1-B5C7-0000F8051515"),Array("Internet Explorer Classes for Java","08B0E5C0-4FCB-11CF-AAA5-00401C608555"),Array("Internet Explorer Help","45EA75A0-A269-11D1-B5BF-0000F8051515"),Array("Internet Explorer Help Engine","DE5AED00-A4BF-11D1-9948-00C04F98BBC9"),Array("NetMeeting NT","44BBA842-CC51-11CF-AAFA-00AA00B6015B"),Array("Task Scheduler","CC2A9BA0-3BDD-11D0-821E-444553540000"),Array("VRML 2.0 Viewer","90A7533D-88FE-11D0-9DBE-0000C0411FC3"),Array("Wallet","1CDEE860-E95B-11CF-B1B0-00AA00BBAD66"));try{aClientInfo=Array(Array("JavaApplet","",(navigator.javaEnabled()?"��":"��"),1),Array("�ͻ��˵�ַ","Client Address",'<%=Request.ServerVariables("REMOTE_ADDR")%>',1),Array("�ͻ��˶˿ں�","Client Port",'<%=Request.ServerVariables("REMOTE_PORT")%>',1),Array("CPU����","CPU Class",navigator.cpuClass,1),Array("����ϵͳ","Operating System",getOS(userAgent),1),Array("����","System Language",(navigator.language?navigator.language:navigator.browserLanguage),1),Array("�����","Browser",getBrower(userAgent),1),Array("�����ȫ��","Browser Name",navigator.appName,1),Array("���������","Browser Code",navigator.appCodeName+" "+parseFloat(navigator.appVersion),1),Array("�����汾","Browser Minor Version",navigator.appMinorVersion,1),Array("MSN Broswer","",getMSN(userAgent),1),Array(".NET CLR","",getDotNet(userAgent),1),Array("JavaScript","JavaScript Support",(bJavaScript ? "�� ( "+getEngVerJs()+" )" :"��"),1),Array("VBScript","Visual Basic Scripting Support",(bVBScript ? "�� ( "+getEngVerVBS()+" )" :"��"),1),Array("Cookies","",(navigator.cookieEnabled?"��":"��"),1));}catch(e){alert(e)}for(var i=0;i<aPlug.length;i++) checkPlug(aPlug [i][0],aPlug [i][1],1);addRow("tblClientInfo",aClientInfo,2);if(!getObjByID("chkAspSupport")){document.title=CCNS_program+" v"+ CCNS_version ;oDiv=getObjByID("divServer");if(oDiv){oDiv.style.display="none";oDiv.id="DisabledDiv";}oDiv=getObjByID("divNoASP");if(oDiv){oDiv.style.display="";oDiv.id="divServer";}}if(oDiv=getObjByID("divStatus")) oDiv.innerHTML="<font color=blue style='border:solid 1px blue;padding-top:1px;padding-left:5px;padding-right:5px;'>������Ϣ</font>��л��ʹ��COCOON ASPϵͳ����̽�롣  [������������ر�̽��]";}function showInfo(iParam){var oServer=getObjByID("divServer");var oClient=getObjByID("divClient");var oFlag=getObjByID("divFlag");if(!oServer||!oClient) return false;if(iParam==1){oFlag.className="tr2";oFlag.innerText="C";oServer.style.display="none";oClient.style.display="";}else{oFlag.className="tr1";oFlag.innerText="S";oServer.style.display="";oClient.style.display="none";}}function runTest(o){var bTest, tTimeBegin, tTimeEnd, iResultTime, i, oElement;tTimeBegin=new Date();bTest=true;iResultTime=0;for(i=0;i<5000000;i++);tTimeEnd=new Date();iResultTime=(tTimeEnd - tTimeBegin) / 1000;o.innerHTML=iResultTime+" ��. ��"+ ((Math.round(5000000/iResultTime*100))/100)+" ��/��.";}</script><script language="JavaScript">
bJavaScript=true;</script><script language="VBScript">bVBScript=true</script></head><script language="javascript">
function showSingleTable(o){var p=o;while(p&&'DIV,BODY'.indexOf(p.parentElement.tagName)<0) p=p.parentElement;var sHtml=p.outerHTML;var sCss="<style>"+document.styleSheets [0].cssText+"</style>";var sOut="<html><head><title>COCOON ASP̽��</title>"+sCss+"</head><body style='overflow:auto;border:0px;background-color:buttonface;text-align:center'>"+sHtml+"</body></html>";sOut="<"+"script>function showSingleTable(){}\nfunction showTools(){}\n</script"+">\n"+ sOut;var oNewWin=window.open("about:blank","","width=790px,height=10px,resizable=yes,scrollbars=yes");oNewWin.document.write(sOut);oNewWin.resizeTo(790,oNewWin.document.body.scrollHeight+35);oNewWin.focus();}function showTools(){document.write('<label onclick="showSingleTable(this)" style="font-family:webdings;color:white;cursor:hand;" title="������ʾ����">2</label> '+'<a href="#top" style="font-family:webdings;color:white;">5</a>');}</script><script language="JavaScript">
var CCNS_program="COCOON ASP ̽��";var CCNS_version="2.9.00";document.title=CCNS_program+" v"+ CCNS_version+"<%=' on '+ sServerName.toUpperCase()+ '  [ '+ sServerAddr+ ' ]'%>";</script><body><a name="top"></a><table width="750" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td width="230" style="font-family:Verdana, Arial, Helvetica">
		<p style="margin-top:0;margin-bottom:-5;font-size:8pt"><strong>COCOON</strong></p>
		<p style="margin-top:0;margin-bottom:-8;">&nbsp;<strong style="font-size:24pt">ASP 
		̽��</strong><font color="#666666">v<script language="JavaScript">document.write(CCNS_version)</script></font></p>
		<p style="margin-top:0;">&nbsp;<font color="#333333" style="font-size:9pt"><u>Server/Client 
		Environment Probe</u></font></p>
		</td>
		<td align="center">
		<table border="0" cellpadding="0" cellspacing="0" style="width:480;height:60;border:solid 1px black;text-align:left;padding-left:10pt">
			<tr>
				<td height="56"><span id="divCcAd" name="divCcAd" style>* ��ӭʹ��COCOON̽�룬�����򹫿�Դ���룬��������⸴�ơ�������ʹ�á�<br>
&nbsp; ����Դ����ǵ���վ ( <a href="http:," www.ccopus.com">www.ccopus.com</a>) ������֧��վ�����ص�������</span></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<% Response.Flush() %><table width="750" border="0" cellspacing="0" cellpadding="0" style="width:750;border:0px solid black;padding:3px;padding-right:0px">
	<tr>
		<td>
		<div id="divStatus" style="font-family:Verdana;cursor:hand;" onclick="window.open('http:," www.ccopus.com/code/aspsyscheck.asp')">
			<strong style="font-size:16px;color:red">Loading...</strong></div>
		</td>
		<td align="right" width="204" style="padding-right:3px"><nobr>
		<a href="javascript:void(0);" class="td1o2" onfocus="this.blur();" onclick="showInfo(0)">
		�� Server Side</a>
		<a href="javascript:void(0);" class="td2o2" onfocus="this.blur();" onclick="showInfo(1)">
		�� Client Side</a></nobr></td>
		<td width="22" style="text-align:right">
		<div class="tr1" id="divFlag" style="width:18px;height:28px;font-family:Times New Roman;font-size:18px;color:#ffffff;text-align:center;padding-top:3px;">
			S</div>
		</td>
	</tr>
</table>
<a name="ServerInfo"></a><a name="ClientInfo"></a>
<table width="750" border="0" cellspacing="0" cellpadding="0" style="width:750;border:0px solid black;padding:3px;padding-right:0px">
	<tr>
		<td align="left"><nobr>����������<a href="#ServerInfo" onclick="javascript:showInfo(0);">&#9674; 
		������Ϣ</a> <a href="#ServerComponent" onclick="javascript:showInfo(0);">&#9674; 
		�����Ϣ</a> <a href="#ArithmeticTest" onclick="javascript:showInfo(0);">&#9674; ��������</a>
		<a href="<%=sTheFile+'?'+Math.random()%>&SpeedTest#SpeedTest" class="td102a">
		&#9830; �����ٶ�</a>
		<a href="<%=sTheFile+'?DriverInfo#DriverInfo'%>" class="td102a">&#9830; ������Ϣ</a>
		<a href="<%=sTheFile+'?ServerDetail#ServerDetail'%>" class="td102a">&#9830; ��������</a>
		<a href="<%=sTheFile+'?DatabaseDriver#DatabaseDriver'%>" class="td102a">
		<font color="#DD5B14">&#9830; ���ݿ�����</font></a> -���ͻ��ˡ�<a href="#ClientInfo" onclick="javascript:showInfo(1);">&#9674; 
		������Ϣ</a> <a href="#ClientComponent" onclick="javascript:showInfo(1);">&#9674; 
		�����Ϣ</a>
		<a href="#ClientArithmeticTest" onclick="javascript:showInfo(1);">&#9674; ��������</a>
		</nobr></td>
	</tr>
</table>
<% Response.Flush() %><div id="divServer">
	<table width="750" border="0" cellpadding="3" cellspacing="1" class="tbl1">
		<tr>
			<td class="tr1">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><font color="#FFFFFF" face="webdings">8</font>
					<font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>�������˻�����Ϣ </strong><strong>
					<script>showTools();</script>
					:::... </strong></font></td>
					<td align="right"><font color="#D2D8EC">Coding by Sunrise_Chen.</font>
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F8F9FC">
			<table width="100%" border="0" cellpadding="3" cellspacing="1" class="tbl1o1">
				<tr>
					<td width="130" class="td1"><font color="#5C72BA">&nbsp;</font>��Ŀ</td>
					<td colspan="3" class="td1"><font color="#5C72BA">&nbsp;</font>ֵ</td>
				</tr>
				<tr>
					<td width="130" class="td1o1">&nbsp;����<br>
					<font color="#666666">&nbsp;Domain Name</font></td>
					<td colspan="3" class="td1o2">&nbsp;<%=Request.ServerVariables("SERVER_NAME")%> 
					/ <%=Request.ServerVariables("LOCAL_ADDR")%></td>
				</tr>
				<tr>
					<td width="130" class="td1o1">&nbsp;�������˿�<br>
					<font color="#666666">&nbsp;Server Port</font></td>
					<td width="240" class="td1o2">&nbsp;<%=Request.ServerVariables("SERVER_PORT")%></td>
					<td width="130" class="td1o1">&nbsp;�ű���ʱʱ��<br>
					<font color="#666666">&nbsp;Script Timeout</font></td>
					<td width="240" class="td1o2">&nbsp;<%=iServerTimeout%>��</td>
				</tr>
				<tr>
					<td width="130" class="td1o1">&nbsp;��Ϣ�������<br>
					<font color="#666666">&nbsp;Server Software</font></td>
					<td width="240" class="td1o2">&nbsp;<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
					<td width="130" nowrap class="td1o1">&nbsp;CPU����<br>
					<font color="#666666">&nbsp;Number of Processors</font></td>
					<td width="240" class="td1o2">&nbsp;<%=isNaN(sNumberOfCpu)?' [δ֪]':sNumberOfCpu+' ��'%></td>
				</tr>
				<tr>
					<td width="130" class="td1o1">&nbsp;����������ϵͳ<br>
					<font color="#666666">&nbsp;Operating System</font></td>
					<td width="240" class="td1o2">&nbsp;<%=sOsInfo%></td>
					<td width="130" class="td1o1">&nbsp;ϵͳ�ļ���<br>
					<font color="#666666">&nbsp;System Path</font></td>
					<td width="240" class="td1o2"><%if(sOsInfo.indexOf("Window")>-1) Response.Write(" &nbsp;"+Request.ServerVariables("windir")+'');else Response.Write(" &nbsp; [δ֪]");%></td>
				</tr>
				<tr>
					<td width="130" nowrap class="td1o1">&nbsp;������������<br>
					<font color="#666666">&nbsp;Application Count</font></td>
					<td width="240" nowrap class="td1o2">&nbsp;<%=Application.Contents.Count%>�� 
					[<a href="<%=sUrl%>?Application#Application" class="td102a">����Application����</a>]
					</td>
					<td nowrap class="td1o1">&nbsp;�Ự������<br>
					<font color="#666666">&nbsp;Session Count</font></td>
					<td width="240" class="td1o2">&nbsp;<%=Session.Contents.Count%>�� 
					<span class="td1o21">[<a href="<%=sUrl%>?Session#Session" class="td102a">����Session����</a>]
					</span></td>
				</tr>
				<tr>
					<td width="130" class="td1o1">&nbsp;����·��<br>
					<font color="#666666">&nbsp;Full path &amp;url</font></td>
					<td colspan="3" class="td1o2">&nbsp;URL:<%var sHostName=Request.ServerVariables("HTTP_HOST");var sPostNo=Request.ServerVariables("SERVER_PORT");var sUrl=Request.ServerVariables("URL");Response.Write("http://"+ sHostName+ (sPostNo=='80'?'':':'+sPostNo)+ sUrl);%><br>
&nbsp;Path:<%=Request.ServerVariables("PATH_TRANSLATED")%></td>
				</tr>
				<tr>
					<td width="130" class="td1o1">&nbsp;�ű�����<br>
					<font color="#666666">&nbsp;ScriptEngine</font></td>
					<td width="240" class="td1o2">&nbsp;JScript / <% try{Response.Write(getEngVerJs())}catch(e){} %>| 
					VBScript / <% Response.Write(getEngVerVBS()) %></td>
					<td nowrap class="td1o1">&nbsp;��ǰ�Ự���<br>
					<font color="#666666">&nbsp;Session Id</font></td>
					<td width="240" class="td1o2">&nbsp;<%=Session.SessionID%></td>
				</tr>
				<tr>
					<td width="130" class="td1o1">&nbsp;��ǰʱ��<br>
					<font color="#666666">&nbsp;Current Time</font></td>
					<td colspan="3" class="td1o2">&nbsp;������:<%=new Date()%><br>
&nbsp;�ͻ���:<script language="JavaScript">document.write(new Date())</script></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td height="5" class="tr1"></td>
		</tr>
	</table>
	<% Server.ScriptTimeout=1800 %><% Response.Flush() %><a name="Application"></a><% if(Request.QueryString("Application").Count>0){ %><br>
	<table width="750" border="0" cellpadding="3" cellspacing="1" class="tbl1">
		<tr>
			<td class="tr1">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><font color="#FFFFFF" face="webdings">8</font>
					<font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>����������Application���� [�� <%=Application.Contents.Count%>��]<script>
					showTools();</script>:::...</strong></font></td>
					<td align="right"><font color="#D2D8EC">Coding by Sunrise_Chen.
					</font></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F8F9FC">
			<table width="100%" border="0" cellpadding="3" cellspacing="1" class="tbl1o1">
				<tr>
					<td width="35%" class="td1"><font color="#5C72BA">&nbsp;����</font></td>
					<td width="65%" class="td1">&nbsp;ֵ</td>
				</tr>
				<%var iCount=0;var sVarType="";var oApplication=new Enumerator(Application.Contents);var oApp;for(;!oApplication.atEnd();oApplication.moveNext()){oApp=oApplication.item();sVarType=typeof(Application.Contents(oApp));++iCount;%><tr>
					<td align="left" class="td1o21"><%=oApp%><i disabled> [<%=sVarType%>]</i></td>
					<td valign="top" class="td1o21"><%try{if(sVarType=="unknown"){var oTmp=new VBArray(Application.Contents(oApp));Response.Write(Server.HTMLEncode(oTmp.toArray()));}else Response.Write(Application.Contents(oApp));}catch(e){Response.Write("<i disabled> [Unknow]</i>");}%></td>
				</tr>
				<%}if(!iCount){%><tr>
					<td align="left" valign="top" class="td1o21" colspan="2">û��Application����</td>
				</tr>
				<%}%></table>
			</td>
		</tr>
		<tr>
			<td height="5" class="tr1" colspan="2"></td>
		</tr>
	</table>
	<% } %><% Response.Flush() %><a name="Session"></a><% if(Request.QueryString("Session").Count>0){ %><br>
	<table width="750" border="0" cellpadding="3" cellspacing="1" class="tbl1">
		<tr>
			<td class="tr1">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><font color="#FFFFFF" face="webdings">8</font>
					<font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>����������Session���� [�� <%=Session.Contents.Count%>��]<script>
					showTools();</script>:::...</strong></font></td>
					<td align="right"><font color="#D2D8EC">Coding by Sunrise_Chen.
					</font></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F8F9FC">
			<table width="100%" border="0" cellpadding="3" cellspacing="1" class="tbl1o1">
				<tr>
					<td width="30%" class="td1"><font color="#5C72BA">&nbsp;����</font></td>
					<td width="70%" class="td1">&nbsp;ֵ</td>
				</tr>
				<%var iCount=0;var sVarType="";var oSession=new Enumerator(Session.Contents);var oSes;for(;!oSession.atEnd();oSession.moveNext()){oSes=oSession.item();sVarType=typeof(Session.Contents(oSes));++iCount;%><tr>
					<td align="left" valign="top" class="td1o21"><%=oSes%><i disabled> 
					[<%=sVarType%>]</i></td>
					<td valign="top" class="td1o21"><%try{if(sVarType=="unknown"){var oTmp=new VBArray(Session.Contents(oSes));Response.Write(Server.HTMLEncode(oTmp.toArray()));}else Response.Write(Session.Contents(oSes));}catch(e){Response.Write("<i disabled> [Unknow]</i>");}%></td>
				</tr>
				<%}if(!iCount){%><tr>
					<td align="left" valign="top" class="td1o21" colspan="2">û��Session����</td>
				</tr>
				<%}%></table>
			</td>
		</tr>
		<tr>
			<td height="5" class="tr1" colspan="2"></td>
		</tr>
	</table>
	<% } %><% Response.Flush() %><% if(Request.QueryString("DriverInfo").Count>0){ %><br>
	<div id="divDriverInfoLoading">
		<font face="Courier New" id="divDriverInfoLoadingAction" style="width:5px;">
		-</font>���ڶ�ȡ������������Ϣ�����ܻỨ��һЩʱ�䣬�����ĵȴ�...<script language="JavaScript">showLoadingAnimation('divDriverInfoLoadingAction')</script></div>
	<% Response.Flush() %><a name="DriverInfo"></a><table width="750" border="0" cellpadding="3" cellspacing="1" class="tbl1">
		<tr>
			<td class="tr1">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><font color="#FFFFFF" face="webdings">8</font>
					<font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>������������Ϣ 
					<script>showTools();</script>
					:::...</strong></font></td>
					<td align="right"><font color="#D2D8EC">Coding by Sunrise_Chen.
					</font></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F8F9FC">
			<table width="100%" border="0" cellpadding="3" cellspacing="1" class="tbl1o1">
				<tr>
					<td width="30%" class="td1"><font color="#5C72BA">&nbsp;��ǰ�ļ���</font></td>
					<td width="70%" class="td1">&nbsp;������Ϣ</td>
				</tr>
				<%function getStrType(s){switch(s){case 0:return "Unknown";case 1:return "3.5 ����";case 2:return "����Ӳ��";case 3:return "Network";case 4:return "CD ������";case 5:return "RamDisk";}}function getSize(iBytes){if(iBytes<1024) return iBytes+' bytes';else if(iBytes/1024<1024) return Math.round(iBytes/1024*100)/100+ ' KB';else if(iBytes/1024/1024<1024) return Math.round(iBytes/1024/1024*100)/100+ ' MB';else return Math.round(iBytes/1024/1024/1024*100)/100+ ' GB';}var bFso=true;try{var oFso=Server.CreateObject("Scripting.FileSystemObject");%><tr>
					<td align="left" valign="top" class="td1o2"><%try{var sCurrPath=Server.MapPath(".");var oFolder=oFso.GetFolder(sCurrPath);var sOut="<div>��ǰ�ļ��У�"+sCurrPath+"</div>"+"<div>���ļ�������"+oFolder.SubFolders.Count+"���ļ�����"+oFolder.Files.Count+"</div>"+"<div>��С��"+getSize(oFolder.Size)+"</div>";Response.Write(sOut);}catch(e){Response.Write(e.description);}%></td>
					<td valign="top" class="td1o2" style="padding:7px;">
					<table width="100%" border="0" cellspacing="1" cellpadding="0">
						<%try{var oDrivers=oFso.Drives;var sOut="<tr align='right' bgColor='#dddddd'>"+"<td style='text-align:center'>����</td>"+"<td style='width:100px'>����&nbsp;</td>"+"<td style='width:100px'>�ļ�ϵͳ&nbsp;</td>"+"<td style='width:100px;'>���ÿռ�&nbsp;</td>"+"<td style='width:100px;'>������&nbsp;</td>"+"</tr>";Response.Write(sOut);for(var x=new Enumerator(oDrivers);!x.atEnd();x.moveNext()) {var oDriver=x.item();var sOut="<tr align='right'>"+"<td align='center'>"+oDriver.Path+"</td>"+"<td>"+getStrType(oDriver.DriveType)+"&nbsp;</td>"+"<td>"+(oDriver.isReady?oDriver.FileSystem:'N/A')+"&nbsp;</td>"+"<td>"+(oDriver.isReady?getSize(oDriver.AvailableSpace):'N/A')+"&nbsp;</td>"+"<td>"+(oDriver.isReady?getSize(oDriver.TotalSize):'N/A')+"&nbsp;</td>"+"</tr>";Response.Write(sOut);}}catch(e){Response.Write(e.description);}%></table>
					</td>
				</tr>
				<%}catch(e){Response.Write("<tr><td colspan=2 class=td1o2>������ �˲�����Ҫ������֧��FileSystemObject�������ķ�������֧�ָ����������������������Ϣ����</span></tr>")}%></table>
			</td>
		</tr>
		<tr>
			<td height="5" class="tr1"></td>
		</tr>
	</table>
	<script language="JavaScript">document.getElementById('divDriverInfoLoading').style.display='none';</script>
	<%}Response.Flush();%><% if(Request.QueryString("ServerDetail").Count>0){ %><a name="ServerDetail"></a><br>
	<table width="750" border="0" cellpadding="3" cellspacing="1" class="tbl1">
		<tr>
			<td class="tr1">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><font color="#FFFFFF" face="webdings">8</font>
					<font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>�������˻��������б� 
					<script>showTools();</script>
					:::...</strong></font></td>
					<td align="right"><font color="#D2D8EC">Coding by Sunrise_Chen.
					</font></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F8F9FC">
			<table width="100%" border="0" cellpadding="3" cellspacing="1" class="tbl1o1">
				<tr>
					<td width="30%" class="td1"><font color="#5C72BA">&nbsp;</font>��Ŀ</td>
					<td class="td1"><font color="#5C72BA">&nbsp;</font>ֵ</td>
				</tr>
				<%try{var WshShell=Server.CreateObject("WScript.Shell");var WshSysEnv=new Enumerator(WshShell.Environment("SYSTEM"));for(;!WshSysEnv.atEnd();WshSysEnv.moveNext()) { %><tr>
					<td class="td1o1">&nbsp;<%=WshSysEnv.item().split("=") [0]%></td>
					<td class="td1o2" style="padding:7px;"><%=WshSysEnv.item().split("=") [1]%></td>
				</tr>
				<%}}catch(e){}%><% for(var oSV=new Enumerator(Request.ServerVariables);!oSV.atEnd();oSV.moveNext()) { x=oSV.item();%><tr>
					<td class="td1o1">&nbsp;<%=x%></td>
					<td class="td1o2" style="padding:7px;"><%=Request.ServerVariables(x).Item.replace(/[\r\n]/g,'<br>')%></td>
				</tr>
				<% } %></table>
			</td>
		</tr>
		<tr>
			<td height="5" class="tr1"></td>
		</tr>
	</table>
	<% Response.Flush() %><% } %><a name="ServerComponent"></a><br>
	<table width="750" border="0" cellpadding="3" cellspacing="1" class="tbl1">
		<tr>
			<td class="tr1">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><font color="#FFFFFF" face="webdings">8</font>
					<font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>�����������Ϣ 
					<script>showTools();</script>
					:::...</strong></font></td>
					<td align="right"><font color="#D2D8EC">Coding by Sunrise_Chen.
					</font></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F8F9FC">
			<table width="100%" border="0" cellpadding="3" cellspacing="1" class="tbl1o1">
				<tr>
					<td width="80%" class="td1">&nbsp;�����</td>
					<td width="10%" align="center" class="td1">֧��</td>
					<td width="10%" align="center" class="td1">�汾</td>
				</tr>
				<%var iIndex, sObjName, sObjDetail, sObjType, iIndexCount, bShowUnSupport;var sObjVersion,bObjInstalled;iIndex=0;iIndexCount=-1;bShowUnSupport=false;for(var i=0;;i++){if(i>=arrObj.length||iIndex!=arrObj [i][2]){if(iIndexCount==0&&iIndex<99){Response.Write('<tr>'+'<td width="80%" class="td1o1" style="color:#666666" colspan="3">&nbsp;<i>( �������ݲ�֧�ָ������ )</i></td>'+'</tr>');}if(i>=arrObj.length) break;iIndex=arrObj [i][2];iIndexCount=0;switch(iIndex){case 1:sObjType="IIS�Դ����";break;case 2:sObjType="�����ĵ������ļ��ϴ��͹������";break;case 3:sObjType="�����ĵ������ļ��ʼ��������";break;case 4:sObjType="������ͼ�������";break;default:sObjType="�Զ������";break;}Response.Write('<tr>'+'<td colspan="3" bgcolor="#D6DBED">&nbsp;��<b>'+sObjType+'</b></td>'+'</tr>');} sObjName=arrObj [i][0];sObjDetail=arrObj [i][1];bObjInstalled=true;sObjVersion=getObjVer(sObjName);if(sObjVersion=="N/A"){bObjInstalled=false;sObjVersion="";}if(bObjInstalled||iIndex==99){iIndexCount++;Response.Write('<tr>'+'<td width="80%" class="td1o1">'+' &nbsp;'+ sObjName+ (sObjDetail?' <font color="#666666">( '+ sObjDetail+ ' )</font>':'')+ ''+'</td>'+'<td width="10%" align="center" class="td1o2">'+ (bObjInstalled?"��":"��")+ '</td>'+'<td width="10%" align="center" class="td1o2"><span style="width:100%;height:18;overflow-y:auto;">'+ sObjVersion+ '</span></td>'+'</tr>');}} %><tr>
					<form method="post" action="<%=sTheFile%>">
						<td colspan="3" bgcolor="#D6DBED">&nbsp;���������� 
						<input name="sObjName" type="text" class="tbl1" id="sObjName" style="background-color:#F3F4FA;width:400px"><input name="Submit" type="submit" value="�ύ" style="border:1px;background-color:#336699;color:#d2d8ec;padding-top:1px;width:70px"></td>
					</form>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td height="5" class="tr1"></td>
		</tr>
	</table>
	<% Response.Flush() %><% if(Request.QueryString("DatabaseDriver").Count>0){ %><br>
	<div id="divDriverInfoLoading">
		<font face="Courier New" id="divDriverInfoLoadingAction" style="width:5px;">
		-</font>���ڶ�ȡ����������������Ϣ�����ܻỨ��һЩʱ�䣬�����ĵȴ�...<script language="JavaScript">showLoadingAnimation('divDriverInfoLoadingAction')</script></div>
	<% Response.Flush() %><a name="DatabaseDriver"></a><table width="750" border="0" cellpadding="3" cellspacing="1" class="tbl1">
		<tr>
			<td class="tr1">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><font color="#FFFFFF" face="webdings">8</font>
					<font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>���ݿ��������� (ODBC)
					<script>showTools();</script>
					:::...</strong></font></td>
					<td align="right"><font color="#D2D8EC">Coding by Sunrise_Chen.
					</font></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F8F9FC">
			<table width="100%" border="0" cellpadding="3" cellspacing="1" class="tbl1o1">
				<tr>
					<td width="35%" class="td1"><font color="#5C72BA">&nbsp;����������</font></td>
					<td width="5%" class="td1" align="center">&nbsp;֧��</td>
					<td width="35%" class="td1"><font color="#5C72BA">&nbsp;����������</font></td>
					<td width="5%" class="td1" align="center">&nbsp;֧��</td>
				</tr>
				<%var aOdbc=new Array("Microsoft Access Driver (*.mdb)","Microsoft Excel Driver (*.xls)"	,"Microsoft Paradox Driver (*.db )","Microsoft Text Driver (*.txt;*.csv)","Microsoft dBASE Driver (*.dbf)","Microsoft Visual FoxPro Driver","SQL Server","Microsoft ODBC for Oracle","Oracle ODBC Driver","Sybase System 11","SYBASE ASE ODBC Driver","Sybase SQL Anywhere 5.0","INTERSOLV 3.10 32-BIT Sybase","MySQL","MySQL ODBC 3.51 Driver","Lotus NotesSQL 3.01 (32-bit) ODBC DRIVER (*.nsf)","MIMER","INFORMIX 3.30 32 BIT","Informix-CLI 2.5 (32 Bit)","Easysoft IB6 ODBC","INTERSOLV InterBase ODBC Driver (*.gdb)","Client Access ODBC Driver (32-bit)");for(var i=0;i<aOdbc.length;++i){var sSQLState="0";try{var oConn=new ActiveXObject("Adodb.Connection");oConn.ConnectionString="Driver={"+aOdbc [i]+"}";oConn.ConnectionTimeout=1;oConn.Open();}catch(e){sSQLState=oConn.Errors(0).SQLState;}if(i%2==0) Response.write("<tr>")%><tr>
					<td align="left" class="td1o1">&nbsp; <%=aOdbc [i]%></td>
					<td class="td1o21" align="center"><%=sSQLState=="IM002"?"��":"��"%></td>
					<%if(i%2==1) Response.write("</tr>");}%></tr>
			</table>
			</td>
		</tr>
		<tr>
			<td height="5" class="tr1" colspan="2"></td>
		</tr>
	</table>
	<script language="JavaScript">document.getElementById('divDriverInfoLoading').style.display='none';</script>
	<%}%><%Server.ScriptTimeout=360;var bTest, tTimeBegin, tTimeEnd, iResultTime1,iResultTime2,sResultTime1,sResultTime2, i;tTimeBegin=new Date();for(i=0;i<500000;++i);tTimeEnd=new Date();iResultTime1=(tTimeEnd - tTimeBegin) / 1000;sResultTime1=iResultTime1+" ��.  ��"+((Math.round(500000/iResultTime1*100))/100)+" ��/��.";tTimeBegin=new Date();for(i=0;i<200000;++i) Math.sqrt(2);tTimeEnd=new Date();iResultTime2=(tTimeEnd - tTimeBegin) / 1000;sResultTime2=iResultTime2+" ��.  ��"+ ((Math.round(200000/iResultTime2*100))/100)+" ��/��."%><br>
	<a name="ArithmeticTest"></a>
	<table width="750" border="0" cellpadding="3" cellspacing="1" class="tbl1">
		<tr>
			<td class="tr1">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><font color="#FFFFFF" face="webdings">8</font>
					<font color="#FFFFFF"><strong>�������������� </strong>(
					<span style="cursor:hand" title="����:50��ε������㲢��ֵ">50���&quot;�ӷ�&quot;����</span>&amp;<span style="cursor:hand" title="����:20��ε���,��ֵ�Լ�20���2��2�η�������">20���&quot;����&quot;����</span>)
					</font>
					<font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>
					<script>showTools();</script>
					</strong></font><font color="#FFFFFF"><strong>:::...</strong></font></td>
					<td align="right"><font color="#D2D8EC">Coding by Sunrise_Chen.
					</font></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F8F9FC">
			<table width="100%" border="0" cellpadding="3" cellspacing="1" class="tbl1o1">
				<tr>
					<td class="td1">&nbsp;������</td>
					<td width="200" align="center" class="td1">����������</td>
					<td width="200" align="center" class="td1">����������</td>
				</tr>
				<tr>
					<td class="td1o1">&nbsp;<a href="http://www.texindex.com.cn/">�л���֯������<em><font color="#666666">(TEXINDEX.com)</font></em></a> 
					[ 2003/6/18 15:08 ]</td>
					<td align="center" class="td1o2">1.297 ��. ��385505.01 ��/��.</td>
					<td align="center" class="td1o2">1.422 ��. ��140646.98 ��/��.</td>
				</tr>
				<tr>
					<td class="td1o1">&nbsp;<a href="http://www.jicang.com/">�й�����������<em><font color="#666666">(JICANG.com)</font></em></a> 
					[ 2003/6/18 15:04 ]</td>
					<td align="center" class="td1o2">0.516 ��. ��968992.25 ��/��.</td>
					<td align="center" class="td1o2">0.609 ��. ��328407.22 ��/��.</td>
				</tr>
				<tr>
					<td class="td1o1">&nbsp;<a href="http://www.ccopus.com">COCOON 
					STUDIO</a><em><font color="#666666">(CCopus.com)</font></em> 
					[ 2003/9/4 11:23 ]</td>
					<td align="center" class="td1o2">0.422 ��. ��1184834.1 ��/��.</td>
					<td align="center" class="td1o2">0.672 ��. ��297619.05 ��/��.</td>
				</tr>
				<tr>
					<td class="td1o1">&nbsp;<a href="http://www.senye.net/">ʤ����������<em><font color="#666666">(Senye.net)</font></em></a> 
					[ 2003/11/5 16:05 ]</td>
					<td align="center" class="td1o2">0.219 ��. ��2283105.0 ��/��.</td>
					<td align="center" class="td1o2">0.265 ��. ��754716.98 ��/��.</td>
				</tr>
				<tr>
					<td class="td1o1" style="color:#336699;font-weight:normal;">&nbsp;��ǰ������ 
					<%=Request.ServerVariables("SERVER_NAME")%></td>
					<td align="center" class="td1o2">
					<a href="<%=sTheFile+'?'+Math.random()%>#ArithmeticTest" title="���²���" onfocus="this.blur();" style="color:#336699">
					<%=sResultTime1%></a></td>
					<td align="center" class="td1o2">
					<a href="<%=sTheFile+'?'+Math.random()%>#ArithmeticTest" title="���²���" onfocus="this.blur();" style="color:#336699">
					<%=sResultTime2%></a></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td height="5" class="tr1"></td>
		</tr>
	</table>
	<%var bSpeedTest=false; if(sQueryString.indexOf('SpeedTest')>-1){%><script language="JavaScript">
	var tSpeedStart=new Date();</script><%bSpeedTest=true;Response.Write("<!--");for(i=1;i<1000;i++)Response.Write("####################################################################################################");Response.Write("-->");%><script language="JavaScript">
	var tSpeedEnd=new Date();</script><script language="JavaScript">var iSpeedTime=0;if(tSpeedEnd>tSpeedStart) iSpeedTime=(tSpeedEnd - tSpeedStart) / 1000;iKbps=Math.round(100 / iSpeedTime * 8 * 10.02) / 10;var iShowPer=Math.round(iKbps / 3000 * 0.8 * 100) / 100 * 100;if(iShowPer<1) iShowPer=1;else if(iShowPer>83) iShowPer=83;</script><%}%><% if(bSpeedTest){ %><script language="JavaScript">
	bSpeedTest='<%=bSpeedTest.toString().toLowerCase()%>';</script><a name="SpeedTest"></a><br>
	<table width="750" border="0" cellpadding="3" cellspacing="1" class="tbl1">
		<tr>
			<td class="tr1">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><font color="#FFFFFF" face="webdings">8</font>
					<font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>�����ٶ� 
					<script>showTools();</script>
					:::...</strong>( ���Ե�ǰ�����£��������Կͻ��˵���Ӧ�ٶ� )</font></td>
					<td align="right"><font color="#D2D8EC">Coding by Sunrise_Chen.
					</font></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#F8F9FC">
			<table width="100%" border="0" cellpadding="3" cellspacing="1" class="tbl1o1">
				<tr>
					<td width="15%" class="td1">&nbsp;�����豸</td>
					<td width="70%" class="td1">&nbsp;�����ٶ� (����ֵ)</td>
					<td class="td1">�����ٶ� (����ֵ)</td>
				</tr>
				<tr>
					<td class="td1o1">56k Modem</td>
					<td class="td1o1">
					<img height="10" align="absmiddle" class="PicBar" style="width:1%">53.3 
					kbps</td>
					<td class="td1o1">&nbsp;3.6 k/s</td>
				</tr>
				<tr>
					<td class="td1o1">64k ISDN</td>
					<td class="td1o1">
					<img height="10" align="absmiddle" class="PicBar" style="width:2%">64 
					kbps</td>
					<td class="td1o1">&nbsp;8.0 k/s</td>
				</tr>
				<tr>
					<td class="td1o1">512k ADSL</td>
					<td class="td1o1">
					<img height="10" align="absmiddle" class="PicBar" style="width:13%">512 
					kbps</td>
					<td class="td1o1">&nbsp;64 k/s</td>
				</tr>
				<tr>
					<td class="td1o1">1.5M Cable</td>
					<td class="td1o1">
					<img height="10" align="absmiddle" class="PicBar" style="width:40%">1500 
					Mbps</td>
					<td class="td1o1">&nbsp;187.5 k/s</td>
				</tr>
				<tr>
					<td class="td1o1">10M FTTP</td>
					<td class="td1o1">
					<img height="10" align="absmiddle" class="PicBar" style="width:83%">10000 
					kbps</td>
					<td class="td1o1">&nbsp;1250 k/s</td>
				</tr>
				<tr>
					<td class="td1o1" style="color:#336699;font-weight:normal;">&nbsp;��ǰ�����ٶ�</td>
					<td nowrap class="td1o1">
					<script language="JavaScript">if(bSpeedTest=='true'){document.write('<img src="" height="10" align="absmiddle" class="PicBar" style="width:'+iShowPer+'%;background-color:#5B64CA">');document.write('<font color="#5c72ba">'+iKbps+' kbps</font>');}</script>
&nbsp;</td>
					<td class="td1o1">&nbsp;<a href="<%=sTheFile+'?'+Math.random()%>&SpeedTest#SpeedTest" title="���������ٶ�" onfocus="this.blur();" style="color:#5c72ba"><u><script language="JavaScript">
					if(bSpeedTest=='true'){document.write(Math.round(iKbps/8*100)/100+ ' k/s');}else{document.write("��ʼ����");}</script></u></a></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td height="5" class="tr1"></td>
		</tr>
	</table>
	<%}%><% Response.Flush() %></div>
<div id="divClient" style="display:none">
	<table width="750" border="0" cellpadding="3" cellspacing="1" class="tbl2">
		<tr>
			<td class="tr2">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><font color="#FFFFFF" face="webdings">8</font>
					<font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>�ͻ��˻�����Ϣ 
					<script>showTools();</script>
					:::...</strong></font></td>
					<td align="right"><font color="#FFFAF7">Coding by Sunrise_Chen.
					</font></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#FFF7F0">
			<table width="100%" border="0" cellpadding="3" cellspacing="1" class="tbl2o1" id="tblClientInfo">
				<tr>
					<td width="60%" class="td2">&nbsp;��Ŀ</td>
					<td align="center" bgcolor="#ABB6DC" class="td2">ֵ</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td height="5" class="tr2"></td>
		</tr>
	</table>
	<% Response.Flush() %><a name="ClientComponent"></a><br>
	<table width="750" border="0" cellpadding="3" cellspacing="1" class="tbl2">
		<tr>
			<td class="tr2">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><font color="#FFFFFF" face="webdings">8</font>
					<font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>�ͻ��������Ϣ 
					<script>showTools();</script>
					:::...</strong></font></td>
					<td align="right"><font color="#FFFAF7">Coding by Sunrise_Chen.
					</font></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#FFF7F0">
			<table width="100%" border="0" cellpadding="3" cellspacing="1" class="tbl2o1" id="tblClientComponent">
				<tr>
					<td width="60%" bgcolor="#ABB6DC" class="td2">&nbsp;��Ŀ</td>
					<td align="center" class="td2">ֵ</td>
				</tr>
				<tr>
					<form method="post" onsubmit="checkPlug(this.clientObj [0].value,this.clientObj [1].value,2);return false;">
						<td colspan="2" class="td2o1">&nbsp;������������ �����(Netscape����)
						<input class="tbl2" id="clientObj" style="WIDTH:100px;BACKGROUND-COLOR:#DDFDF2" name="clientObj"> 
						���ID(IE����)
						<input class="tbl2" id="clientObj" style="WIDTH:240px;BACKGROUND-COLOR:#DDFDF2" name="clientObj">
						<input style="BORDER-RIGHT:1px;BORDER-TOP:1px;BORDER-LEFT:1px;COLOR:#FFF;BORDER-BOTTOM:1px;BACKGROUND-COLOR:#50A0A0;width:70px;padding-top:1px" type="submit" value="�ύ" name="Submit"></td>
					</form>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td height="5" class="tr2"></td>
		</tr>
	</table>
	<% Response.Flush() %><a name="ClientArithmeticTest"></a><br>
	<table width="750" border="0" cellpadding="3" cellspacing="1" class="tbl2">
		<tr>
			<td class="tr2">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><font color="#FFFFFF" face="webdings">8</font>
					<font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">
					<strong>�ͻ�����������</strong>��500��� <strong>���ӷ�</strong>���� ��<strong><script>
					showTools();</script>:::...</strong></font></td>
					<td align="right"><font color="#FFFAF7">Coding by Sunrise_Chen.
					</font></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td bgcolor="#FFF7F0">
			<table width="100%" border="0" cellpadding="3" cellspacing="1" class="tbl2o1">
				<tr>
					<td width="70%" class="td2">&nbsp;�ͻ���</td>
					<td align="center" class="td2">����ʱ��</td>
				</tr>
				<tr>
					<td width="70%" class="td2o1">&nbsp;Sunrise_Chen���Ƶ��� ��192M C500��</td>
					<td align="center" class="td2o2">5.698 ��. ��877500.88 ��/��.</td>
				</tr>
				<tr>
					<td class="td2o1">&nbsp;�й����������� [ 2002/8/28 16:44 ]</td>
					<td align="center" class="td2o2">9.546 ��. ��523779.59 ��/��.</td>
				</tr>
				<tr>
					<td class="td2o1">&nbsp;����Ͷ���������� [ 2002/8/28 15:50 ]</td>
					<td align="center" class="td2o2">3.484 ��. ��1435132.1 ��/��.</td>
				</tr>
				<tr>
					<td class="td2o1">&nbsp;��ǰ�ͻ���</td>
					<td align="center" class="td2o2"><u>
					<a href="javascript:void(0);" onclick="runTest(this)" onfocus="this.blur();" title="���Ե�ǰ�ͻ���">
					��ʼ����</a></u></td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td height="5" class="tr2"></td>
		</tr>
	</table>
	<% Response.Flush() %><br>
</div>
<table width="750" border="0" cellpadding="3" cellspacing="1" class="tbl1" id="divNoASP" style="display:none">
	<tr>
		<td class="tr1">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td><font color="#FFFFFF" face="webdings">8</font>
				<font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">
				<strong>������Ϣ :::...</strong></font></td>
				<td align="right"><font color="#FFFAF7">Coding by Sunrise_Chen.
				</font></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#F8F9FC">
		<table width="100%" border="0" cellpadding="3" cellspacing="1" class="tbl1o1">
			<tr>
				<td width="60%" class="td1">&nbsp;��Ϣ����</td>
			</tr>
			<tr>
				<td width="60%" class="td1o1">&nbsp;��ǰ��������֧��ASP��</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td height="5" class="tr1"></td>
	</tr>
</table>
<%="<div id='chkAspSupport'></div>"%><script language="JavaScript">pageInit();</script><div align="center" style="width:750px">
	<hr size="1"><a href="http://www.ccopus.com" target="_blank">COCOON sTudio ��Ȩ����</a>��<a href="http://www.ccopus.com/code/aspSysCheck.asp"> 
	[���������ػ���±�̽��]</a><% tPageEndTime=new Date();%><%="<br>ҳ��ִ��ʱ�䣺Լ <font color='#990000'><b>"+((tPageEndTime-tPageStartTime)/1000)+"</b></font>��"%><script langauge="JavaScript" src="http://www.ccopus.com/_js/count_aspcheck.js"></script></div>

</body>

</html>
