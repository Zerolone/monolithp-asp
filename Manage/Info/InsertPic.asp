<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include file ="InfoSetting.asp" -->
<html>

<head>
<script>
function switchCell(n, hash) {
	nc=document.getElementsByName("navcell");
	if(nc){
		t=document.getElementsByName("tb")
		for(i=0;i<nc.length;i++){
			nc.item(i).className="tab-off";
			t.item(i).className="hide-table";
		}
		nc.item(n-1).className="tab-on";
		t.item(n-1).className="show-table";
	}else if(navcell){
		for(i=0;i<navcell.length;i++){
			navcell[i].className="tab-off";
			tb[i].className="hide-table";
		}
		navcell[n-1].className="tab-on";
		tb[n-1].className="show-table";
	}
	if(hash){
		document.location="#"+hash;
	}
}
</script>
<script event="onclick" for="Ok" language="JavaScript">
  window.returnValue = InsertA.value+"*"+InsertB.value+"*"+InsertC.value+"*"+InsertD.value+"*"+InsertE.value+"*"+InsertF.value+"*"+InsertG.value+"*"+InsertH.value+"*"+InsertI.value
  window.close();
</script>

<script event="onclick" for="Ok2" language="JavaScript">
  window.returnValue = "UpLoadPic*" + form1.UpLoadPic.value;
  window.close();
</script>
<script event="onclick" for="Ok3" language="JavaScript">
  window.returnValue = "UpLoadPic*" + UpPicSelectList.value;
  window.close();
</script>
<link href="../../Css/Manage.Css" rel="stylesheet" type="text/css">
</head>

<body>

<table cellspacing="0" cellpadding="0" width="490" border="0">
	<tr>
		<td class="tab-on" id="navcell" onclick="switchCell(1)" name="navcell" width="140" height="30" style="font-size: 12pt; font-weight:bold">
&nbsp;����ͼƬ</td>
		<td class="tab-off" id="navcell" onclick="switchCell(2)" name="navcell" width="140" style="font-size: 12pt; font-weight:bold">
&nbsp;�ϴ�ͼƬ</td>
		<td class="tab-off" id="navcell" onclick="switchCell(3)" name="navcell" width="140" style="font-size: 12pt; font-weight:bold">
&nbsp;�Ѿ��ϴ�ͼƬ</td>
		<td width="*">
&nbsp; </td>
	</tr>
</table>
<table id="tb" cellspacing="0" cellpadding="0" width="490" border="0" name="tb" style="border: 1px solid #C0C0C0">
	<tr>
		<td>
		<table border="0" width="100%" id="table1" cellspacing="0" cellpadding="0">
			<tr>
				<td height="22">
&nbsp; ͼƬ��ַ��<input id="InsertA" value="http://" size="30" class="InputBox" style="width=284"></td>
			</tr>
			<tr>
				<td height="22">
&nbsp; ��ʶ���֣�<input id="InsertB" size="12" class="InputBox" style="width=100">&nbsp;&nbsp;&nbsp; ͼƬ�߿�<input id="InsertI" value="0" size="2" maxlength="2" class="InputBox" style="width=100"></td>
			</tr>
			<tr>
				<td height="22">
&nbsp; ����Ч����<select id="InsertC" class="InputBox" style="width=100">
				<option>��Ӧ��</option>
				<option value="filter:blur(add=1,direction=14,strength=15)">ģ��Ч��</option>
				<option value="filter:gray">�ڰ���Ƭ</option>
				<option value="filter:flipv">�ߵ���ʾ</option>
				<option value="filter:fliph">���ҷ�ת</option>
				<option value="filter:xray">X����Ƭ</option>
				<option value="filter:invert">��ɫ��ת</option>
				</select>&nbsp;&nbsp;&nbsp; ͼƬλ�ã�<select id="InsertH" class="InputBox" style="width=100">
				<option>Ĭ��λ��</option>
				<option value="left">����</option>
				<option value="right">����</option>
				<option value="top">����</option>
				<option value="middle">�в�</option>
				<option value="bottom">�ײ�</option>
				<option value="absmiddle">���Ծ���</option>
				<option value="absbottom">���Եײ�</option>
				<option value="baseline">����</option>
				<option value="texttop">�ı�����</option>
				</select></td>
			</tr>
			<tr>
				<td>
&nbsp; ͼƬ��ȣ�<input id="InsertD" value="200" size="4" maxlength="4" class="InputBox" style="width=100">&nbsp;&nbsp;&nbsp; ͼƬ�߶ȣ�<input id="InsertE" value="200" size="4" maxlength="4" class="InputBox" style="width=100"></td>
			</tr>
			<tr>
				<td>
&nbsp; ���¼�ࣺ<input id="InsertF" value="0" size="4" maxlength="2" class="InputBox" style="width=100">&nbsp;&nbsp;&nbsp; ���Ҽ�ࣺ<input id="InsertG" value="0" size="4" maxlength="2" class="InputBox" style="width=100"> </td>
			</tr>
			<tr>
				<td height="35">
				<p align="center">&nbsp;<input type="submit" value=" ȷ �� (Alt + Y) " id="Ok" class="InputBox" accesskey="Y">&nbsp; 
				<input type="button" value=" ȡ �� (Alt + N) " onclick="window.close();" class="InputBox" accesskey="N"></p>
				</td>
			</tr>
		</table>
		</td>
	</tr>
</table>

<table id="tb" cellspacing="0" cellpadding="0" width="490" border="0" name="tb" class="hide-table" style="border: 1px solid #C0C0C0">
	<tr>
		<td>
		<form name="form1" method="post" action="upfile.asp" target="AAA" enctype="multipart/form-data" style="MARGIN-BOTTOM:0px">
			<input type="hidden" name="act" value="upload"><table border="0" cellspacing="0" cellpadding="0">
				<tr align="left">
					<td>
					<script language="javascript">
	  function setid()
	  {
	  var j;
	  str='<br>';
	  if(!window.form1.upcount.value)
	   window.form1.upcount.value=1;
	   if(window.form1.upcount.value > 10)
	   {
	   	alert("һ���ύ�ĸ���̫�� ���趨��10���ڣ�");
	   	return false;
	   }
	   
 	  for(i=1;i<=window.form1.upcount.value;i++)
 	  {
 	  		if(i<10){j= '0'+ i}
 	  		else{j=i}
	     str+='&nbsp;&nbsp;�ļ�'+j+':<input type="file" name="file'+i+'" class="InputBox" style="width=400"><br>';
//	     str+='�ļ�'+i+':<input type="file" name="file'+i+'" style="width:350" class="InputBox" style="width=350"><br><br>';
		}
	  window.upid.innerHTML=str+'<br>';
	  }
	  </script>
					&nbsp;&nbsp;��Ҫ�ϴ��ĸ��� <input type="text" name="upcount" value="<%=InfoSettingArr(1)%>" class="InputBox" style="width=100"> 
					<input type="button" name="Button" onclick="setid();" value=" �� �� (Alt + S) " class="InputBox" accesskey="S"> </li>
					<br>
					&nbsp;&nbsp;�ϴ���:<input type="text" name="filepath" value="<%=InfoSettingArr(0)%>" class="InputBox" style="width=400">
					</td>
				</tr>
				<tr align="center">
					<td align="left" id="upid"><br>
					&nbsp;&nbsp;�ļ�01:<input type="file" name="file1" class="InputBox" style="width=400"> </td>
				</tr>
				<tr align="center" valign="middle">
					<td height="24">
					<input type="submit" name="Submit" value=" �� �� (Alt + U) " class="InputBox" accesskey="U">&nbsp; 
					<input type="button" value=" ȡ �� (Alt + N) " onclick="window.close();" class="InputBox">
					<input type="hidden" value="" name="UpLoadPic"><input type="button" name="InsertUpLoadPic" style="display:none;" value=" �� �� ͼ Ƭ �� �� �� �� " class="InputBox" id="Ok2">
					</td>
				</tr>
			</table>
		</form>
		</td>
	</tr>
</table>
<table id="tb" cellspacing="0" cellpadding="0" width="490" border="0" name="tb" class="hide-table" style="border: 1px solid #C0C0C0">
	<tr>
		<td height="270">
		<iframe name="I1" src="UpLoadedPic.asp" marginwidth="1" marginheight="1" height="100%" width="100%" border="0" frameborder="0">
		�������֧��Ƕ��ʽ��ܣ�������Ϊ����ʾǶ��ʽ��ܡ�</iframe>
		</td>
	</tr>
	<tr>
	<td>
			<form name="form2" style="MARGIN-BOTTOM:0px">
			<input type="hidden" value="" name="UpPicSelectList">
			<p align="center"><input type="button" name="InsertUpLoadPic0" value=" �� �� ͼ Ƭ �� �� �� �� " class="InputBox" id="Ok3">
			</form>
			</td>
	</tr>
</table>
<iframe name="AAA" src="about:link" width="1" height="1"></iframe>
</body>

</html>