/*
'-----------------------------------------------
'����: -----------------------------------------
'����:20041010----------------------------------
'����:�������----------------------------------
'����:------------------------------------------
'-----------------------------------------------
*/
function fortable()
{

  var arr = showModalDialog("table.htm", "", "dialogWidth:250px; dialogHeight:200px; status:1;help:0");
  if (arr != null){
  var tmp=arr.split("*");
  row=tmp[0];
  col=tmp[1];
  var string;
  string="<table style='WIDTH: "+tmp[2]+"px;' border="+tmp[3]+" bgcolor="+tmp[6]+" cellspacing="+tmp[5]+" cellpadding="+tmp[4]+">";
  for(i=1;i<=row;i++){
  string=string+"<tr>";
  for(j=1;j<=col;j++){
  string=string+"<td>&nbsp;</td>";
  }
  string=string+"</tr>";
  }
  string=string+"</table>";
  content=Monolith_News_Body.document.body.innerHTML;
  content=content+string;
  Monolith_News_Body.document.body.innerHTML=content;
  }
  else Monolith_News_Body.focus();
}

/*
'-----------------------------------------------
'����:Zerolone----------------------------------
'����:20041010----------------------------------
'����:�����ý��--------------------------------
'����:------------------------------------------
'-----------------------------------------------
*/
function Allmedia()
{
  var arr = showModalDialog("Media.htm", "", "dialogWidth:500px; dialogHeight:200px; status:1;help:0");

  if (arr != null){
  var ss;
  ss=arr.split("*")
  Code	= ss[0];
  path	= ss[1];
  row	= ss[2];
  col	= ss[3];
  
   var string;
   if (Code=="Flash")
	{
	   string="<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0' width="+row+" height="+col+"><param name='movie' value="+path+"><param name='quality' value='high'></object>";
	   Monolith_News_Body.document.body.innerHTML += string;
	}
   else if (Code=="MediaPlayer")
	{
	   string="<object classid='clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95' width="+row+" height="+col+"><param name=Filename value="+path+"><param name='BufferingTime' value='5'><param name='AutoSize' value='-1'><param name='AnimationAtStart' value='-1'><param name='AllowChangeDisplaySize' value='-1'><param name='ShowPositionControls' value='0'><param name='TransparentAtStart' value='1'><param name='ShowStatusBar' value='1'></object>";
	   Monolith_News_Body.document.body.innerHTML += string;
	}
   else if (Code=="RealPlayer")
	{
	   string="<object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA' width="+row+" height="+col+"><param name='CONTROLS' value='ImageWindow'><param name='CONSOLE' value='Clip1'><param name='AUTOSTART' value='-1'><param name=src value="+path+"></object><br><object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA'  width="+row+" height=60><param name='CONTROLS' value='ControlPanel,StatusBar'><param name='CONSOLE' value='Clip1'></object>";
	   Monolith_News_Body.document.body.innerHTML += string;
	}
 }
 else Monolith_News_Body.focus();
}

/*
'-----------------------------------------------
'����: -----------------------------------------
'����:20041010----------------------------------
'����:����ͼƬ���ϴ�ͼƬ�������Ѿ��ϴ�ͼƬ------
'����:------------------------------------------
'-----------------------------------------------
*/
function pic()
{
  var arr = showModalDialog("InsertPic.asp", "", "dialogWidth:510px; dialogHeight:392px; status:1;help:0");
  
  if (arr != null)
	  {
	  var ss;
	  ss=arr.split("*");
	  a=ss[0];
	  b=ss[1];
	  c=ss[2];
	  d=ss[3];
	  e=ss[4];
	  f=ss[5];
	  g=ss[6];
	  h=ss[7];
	  i=ss[8];
	  
		if (a=="UpLoadPic")
		{
		  Monolith_News_Body.document.body.innerHTML=Monolith_News_Body.document.body.innerHTML + b;
		}
		else
		{
	  var str1;
	  str1="<img src='"+a+"' alt='"+b+"'"
	  if(d.value!='')str1=str1+"width='"+d+"'"
	  if(e.value!='')str1=str1+"height='"+e+"' "
	  str1=str1+" border='"+i+"' align='"+h+"' vspace='"+f+"' hspace='"+g+"'  style='"+c+"'"
	  str1=str1+">"
//	  content=Monolith_News_Body.document.body.innerHTML;
//	  content=content+str1;
//	  Monolith_News_Body.document.body.innerHTML=content;
	  Monolith_News_Body.document.body.innerHTML+=str1;
	  }
		}
  else Monolith_News_Body.focus();
}

/*
'-----------------------------------------------
'����: -----------------------------------------
'����:20041208----------------------------------
'����:�����Ѿ��ϴ�ͼƬ��ָ����News_Pic1---------
'����:------------------------------------------
'-----------------------------------------------
*/
function Insertpic1()
{
  var arr = showModalDialog("InsertPic1.asp", "", "dialogWidth:510px; dialogHeight:360px; status:1;help:0");
  AddNewsForm.News_Pic1.value=  arr;
  AddNewsForm.News_Pic1.focus();
}

function Insertpic2()
{
  var arr = showModalDialog("InsertPic2.asp", "", "dialogWidth:510px; dialogHeight:360px; status:1;help:0");
  AddNewsForm.News_Pic2.value=  arr;
  AddNewsForm.News_Pic2.focus();
}

/*
'-----------------------------------------------
'����:Zerolone----------------------------------
'����:20041208----------------------------------
'����:Ԥ��ͼƬ----------------------------------
'����:ThePicUrl ͼƬ·��------------------------
'-----------------------------------------------
*/
function View(ThePicUrl)
{
	ViewImg.src=ThePicUrl;
	Layer1.style.visibility = "visible";
//    Layer1.style.posLeft = screen.height/2;
    Layer1.style.posLeft = document.body.scrollLeft + event.clientX + 50;
    Layer1.style.posTop  = document.body.scrollTop  + event.clientY + 50;
}

/*
'-----------------------------------------------
'����:Zerolone----------------------------------
'����:20041208----------------------------------
'����:����Ԥ��ͼƬ------------------------------
'����:------------------------------------------
'-----------------------------------------------
*/
function HiddenLayer()
{
	Layer1.style.visibility = "Hidden";
	ViewImg.src				= "../images/LoadingPic.gif";
}

function setMode(newMode)
{
  bTextMode = newMode;
  var cont;
  if (bTextMode) {
    cleanHtml();
    cleanHtml();

    cont=Monolith_News_Body.document.body.innerHTML;
    Monolith_News_Body.document.body.innerText=cont;
  } else {
    cont=Monolith_News_Body.document.body.innerText;
    Monolith_News_Body.document.body.innerHTML=cont;
  }
  Monolith_News_Body.focus();
}

function cleanHtml()
{
  var fonts = Monolith_News_Body.document.body.all.tags("FONT");
  var curr;
  for (var i = fonts.length - 1; i >= 0; i--) {
    curr = fonts[i];
    if (curr.style.backgroundColor == "#ffffff") curr.outerHTML	= curr.innerHTML;
  }

}

function Page()
{
  var str1;
  str1="[---��ҳ---]";
  Monolith_News_Body.document.body.innerHTML+=str1;
}


function help()
{
    var helpmess;
    helpmess="---------------��д����---------------\r\n\r\n"+
         "1.�벻Ҫ������Σ���ԵĽű���\r\n\r\n"+
         "2.���Ҫ��дԴ���룬��ѡ��\r\n\r\n"+
         "���鿴HTMLԴ������д.\r\n\r\n"+
         "3.��Ҫ���Լ�����,���ܿ�Ч��.\r\n\r\n"+
         "4.�����дjs��������Ҫ�������д.\r\n\r\n";
    alert(helpmess);
}

//Function to format text in the text box
function FormatText(command, option){
	
  	frames.Monolith_News_Body.document.execCommand(command, true, option);
  	frames.Monolith_News_Body.focus();
}

/*
'-----------------------------------------------
'����:Discuz��----------------------------------
'����:------------------------------------------
'����:���д���----------------------------------
'����:------------------------------------------
'-----------------------------------------------
*/
function runCode(obj) {
	var winname = window.open('', "_blank", '');
	winname.document.open('text/html', 'replace');
	winname.document.writeln(obj.value.replace(/t extarea/g, "textarea"));
	winname.document.close();
}





