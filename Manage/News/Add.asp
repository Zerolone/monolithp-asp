<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Loading.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include file ="NewsSetting.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link href="/CSS/Manage.Css" rel="stylesheet" type="text/css">
<title>�ޱ����ĵ�</title>
<%
	Dim News_ID
	News_ID	= Request("News_ID")
	CheckNum News_ID
	
	Dim News_Title, News_Title2, News_Class, News_Author, News_From, News_KeyWord, News_Memo, News_Pic1, News_Pic2, News_FileName, News_Date, News_Template, News_Area, News_Order
	News_Area		= 0
	News_Order	= 999
	
	if News_ID <> "" then
		'------------------0-------------1--------------2-------------3--------------4------------5---------------6------------7------------8------------9----------------10-----------11---------------12-----------13
		Sql = "Select [News_Title], [News_Title2], [News_Class], [News_Author], [News_From], [News_KeyWord], [News_Memo], [News_Pic1], [News_Pic2], [News_FileName], [News_Date], [News_Template], [News_Area], [News_Order] From [Monolith_News] Where [News_ID]=" & News_ID
'		Response.write Sql
'		Response.end
		Rs.Open Sql, Conn, 1, 1
		if Not (Rs.Bof Or Rs.Eof) then
			News_Title			= Rs(0)
			News_Title2			= Rs(1)
			News_Class			= Rs(2)
			News_Author			= Rs(3)
			News_From				= Rs(4)
			News_KeyWord		= Rs(5)
			News_Memo				= Rs(6)
			News_Pic1				= Rs(7)
			News_Pic2				= Rs(8)
			News_FileName		= Rs(9)
			News_Date				= Rs(10)
			News_Template		= Rs(11)
			News_Area				= Rs(12)
			News_Order			= Rs(13)
		end if
		Rs.Close
	end if
%>
<script language="javascript" src="/Js/All.js"></script>
<script language="javascript" src="/Js/News_Edit.js"></script>
<script language="javascript" src="/Js/PopupCalendar.js"></script>
<script language="javascript">
function ShowOrHideTr(TrName)
{
	TrName	= eval(TrName);
	if(TrName.style.display == "none")
	{
		TrName.style.display = "";
	}
	else
	{
		TrName.style.display = "none";
	}
}

function Count_News_Title()
{
	var strlength=0;
	var strlength2=0;
	var TheString;
	TheString = AddNewsForm.News_Title.value;
	for (var i=0;i<TheString.length;i++)
	{
		if(TheString.charCodeAt(i)>=1000)
			strlength += 2;
		else
			strlength += 1;
	}
	
	TheString = AddNewsForm.News_Title2.value;
	for (var i=0;i<TheString.length;i++)
	{
		if(TheString.charCodeAt(i)>=1000)
			strlength2 += 2;
		else
			strlength2 += 1;
	}

	alert("��������ⳤ��Ϊ�� " + strlength +  "����\n\n��������ⳤ��Ϊ�� " + strlength2 +  "����");
}

function FrameHeight(Action)
{
	if (Action== "Add")
	{
		TrForBody.height	= parseInt(TrForBody.height) + 50;
	}
	else
	{
		if (TrForBody.height > 490)
		{
			TrForBody.height	= TrForBody.height - 50;
		}
	}
}

function ResetDefault()
{
	TrForTitle.style.display				= "none";
	TrForTitle2.style.display 			= "none";
	TrForText.style.display 				= "none";
	TrForPic.style.display 					= "none";
	TrForNewsArea.style.display 		= "none";
	TrForEToolBar1.style.display 		= "none";
	TrForEToolBar2.style.display 		= "none";
	
	TrForBody.height = 490;
	
}

var oCalendarChs=new PopupCalendar("oCalendarChs");	//��ʼ���ؼ�ʱ,�����ʵ������:oCalendarChs
oCalendarChs.weekDaySting=new Array("��","һ","��","��","��","��","��");
oCalendarChs.monthSting=new Array("һ��","����","����","����","����","����","����","����","����","ʮ��","ʮһ��","ʮ����");
oCalendarChs.oBtnTodayTitle="����";
oCalendarChs.oBtnCancelTitle="ȡ��";
oCalendarChs.Init();

function SaveToRemote()
{
	AddNewsForm.action 							= "SaveRemoteFile.asp";
	AddNewsForm.target							=	"AAA";
	AddNewsForm.News_Content.value	= frames.Monolith_News_Body.document.body.innerHTML;
	LayerPicToRemote.style.visibility	=	"visible";
	AddNewsForm.submit();
}

function AddNews()
{
  var Check_News_Title		= document.AddNewsForm.News_Title;
  var Check_News_Author		= document.AddNewsForm.News_Author;
  var Check_News_Form			= document.AddNewsForm.News_From;
  var Check_News_KeyWord	= document.AddNewsForm.News_KeyWord;
  
  if (Check_News_Title.value == "")
  {
 	alert("���������±��⣡");
	Check_News_Title.focus();
	return false;
	}

  if (Check_News_Author.value == "")
  {
 	alert("���������ߣ�");
	Check_News_Author.focus();
	return false;
  }

  if (Check_News_Form.value == "")
  {
 	alert("��������Դ��");
	Check_News_Form.focus();
	return false;
  }

  if (Check_News_KeyWord.value == "")
  {
 	alert("�ؼ��ֲ���Ϊ�գ�");
	Check_News_KeyWord.focus();
	return false;
  }

	AddNewsForm.action 							= "Add_News.asp";
	AddNewsForm.target							=	"_self";
	AddNewsForm.News_Content.value	= frames.Monolith_News_Body.document.body.innerHTML;
	AddNewsForm.submit();
}

</script>
<div id="Layer1" style="position:absolute; left:200px; top:17px; width:100px; height:100px; z-index:1; visibility:hidden">
	<table border="1" width="100%" id="table1" cellspacing="0" cellpadding="0" bordercolor="#000000" onclick="HiddenLayer();">
		<tr>
			<td><img src="../images/LoadingPic.gif" name="ViewImg"></td>
		</tr>
	</table>
</div>
<form method="POST" name="AddNewsForm">
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
		<tr>
			<td width="100%">
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td width="560" height="22">���±���:<input type="text" name="News_Title" size="80" class="InputBox" style="width=500" onmouseover="select();" value="<%=News_Title%>"></td>
					<td width="60" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="ShowOrHideTr('TrForTitle2')" bgcolor="#C0C0C0">
					��ʾ������</td>
					<td width="60" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="ShowOrHideTr('TrForTitle')" bgcolor="#C0C0C0">
					��ʾ����</td>
					<td width="60" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="Count_News_Title()" bgcolor="#C0C0C0">
					ͳ���ַ�</td>
					<td width="*">&nbsp; ��</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr id="TrForTitle2" style="display:none;">
			<td width="100%">
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td width="560" height="22">�ڶ�����:<input type="text" name="News_Title2" size="80" class="InputBox" style="width=500" onmouseover="select();" value="<%=News_Title2%>"></td>
					<td width="*">&nbsp; ��</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr id="TrForTitle" style="display:none;">
			<td width="100%" height="21">��&nbsp;&nbsp;&nbsp; ��:<input readonly type="text" name="_Title" size="80" class="InputBox" value="һ�����������߰˾�ʮһ�����������߰˾�ʮһ�����������߰˾�ʮһ�����������߰˾�ʮ" style="background-color: #C0C0C0; width=500"></td>
		</tr>
		<tr id="TrForNewsArea" style="display:none;">
			<td width="100%" height="21">��ʾλ��:<input type="checkbox" name="News_Area" value="1" id="NewsArea_01" <%if right(news_area, 1) >= 1 then response.write "checked"%>><label for="NewsArea_01">��ҳ</label><input type="checkbox" name="News_Area" value="10" id="NewsArea_02" <%if clng(right(news_area, 2)) >= 10 then response.write "checked"%>><label for="NewsArea_02">ͼƬ����</label><input type="checkbox" name="News_Area" value="100" id="NewsArea_03" <%if clng(right(news_area, 3)) >= 100 then response.write "checked"%>><label for="NewsArea_03">��Ŀ��ҳ</label><input type="checkbox" name="News_Area" value="1000" id="NewsArea_04" <%if clng(right(news_area, 4)) >= 1000 then response.write "checked"%>><label for="NewsArea_04">��Ŀ��ҳ�Ƽ�</label><input type="checkbox" name="News_Area" value="10000" id="NewsArea_05" <%if clng(right(news_area, 5)) >= 10000 then response.write "checked"%>><label for="NewsArea_05">��Ŀ��ҳ�Ƽ��̳�</label>&nbsp; 
			��ʾ˳��:<input type="text" name="News_Order" size="80" class="InputBox" style="width=124" onmouseover="select();" value="<%=News_Order%>"></td>
		</tr>
		<tr>
			<td width="100%">
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td width="560" height="22">������Ŀ:<select size="1" name="News_Class" class="InputBox" style="width=200">
					<%
			'---------------------0-----------------1------------------2
			Sql	= "Select [News_Class_ID], [News_Class_Level], [News_Class_Title] From [Monolith_News_Class] Order By [News_Class_Level]"
			Response.write Sql
			Rs.Open Sql, Conn, 1, 1
			Do While Not (Rs.Bof Or Rs.Eof)
			%>
					<option value="<%=Rs(1)%>" <% if rs(1)=news_class then response.write " selected"%>>
					<%LoopNBSP(Len(Rs(1))-2)%><%=Rs(1)%>-<%=Rs(2)%></option>
					<%
				Rs.MoveNext
			Loop
			Rs.Close
			%></select>&nbsp; ����ģ��:
					<select size="1" name="News_Template" class="InputBox" style="width=230">
					<%
			Dim HasTemplate
			HasTemplate = false
			'---------------------0-----------------1------------------2
			Sql	= "Select [Template_ID], [Template_Level], [Template_Title] From [Monolith_Template] Order By [Template_Level]"
			Response.write Sql
			Rs.Open Sql, Conn, 1, 1
			Do While Not (Rs.Bof Or Rs.Eof)
			%>
					<option value="<%=Rs(0)%>" <%
				if rs(0)=news_template then
					response.write " selected" 
					hastemplate		= true
				end if
				if rs(0)=cint(newssettingarr(5)) and hastemplate=false then response.write " selected"
			%>><%LoopNBSP(Len(Rs(1))-2)%><%=Rs(1)%>-<%=Rs(2)%></option>
					<%
				Rs.MoveNext
			Loop
			Rs.Close
			%></select></td>
					<td width="60" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="ShowOrHideTr('TrForText')" title="��ʾ���ص��õ�����" bgcolor="#C0C0C0">
					��������</td>
					<td width="60" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="ShowOrHideTr('TrForPic')" title="��ʾ���ص��õ�ͼƬ" bgcolor="#C0C0C0">
					����ͼƬ</td>
					<td width="60" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="ResetDefault()" title="�༭�����ָ�ΪĬ��" bgcolor="#C0C0C0">
					����ΪĬ��</td>
					<td width="*">&nbsp; ��</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100%">
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td width="560">��&nbsp;&nbsp;&nbsp; ��:<input type="text" name="News_Author" size="30" class="InputBox" style="width=125" onmouseover="select();" value="<%=News_Author%>" onclick="ShowCombo(1);">&nbsp; 
					��&nbsp;&nbsp; Դ:<input type="text" name="News_From" size="31" class="InputBox" style="width=125" onmouseover="select();" value="<%=News_From%>" onclick="ShowCombo(2);">&nbsp; 
					�� �� ��:<input type="text" name="News_KeyWord" size="80" class="InputBox" style="width=124" onmouseover="select();" value="<%=News_KeyWord%>" onclick="ShowCombo(3);">
					</td>
					<script language="javascript">
						function ShowCombo(Num)
						{
							Layer_News.style.visibility="";
							
							switch(Num)
							{
								case 1:
								     AddNewsForm.List_News_Author.style.visibility="";
								     break;
								case 2:
								     AddNewsForm.List_News_From.style.visibility="";
								     break;
								case 3:
								     AddNewsForm.List_News_KeyWord.style.visibility="";
								     break;
								default: 
											//Do Nothing
							}

						} 
						
						function AddText(Num)
						{
							switch(Num)
							{
								case 1:
										 AddNewsForm.News_Author.value	= AddNewsForm.List_News_Author.item(AddNewsForm.List_News_Author.selectedIndex).value;
								     break;
								case 2:
										 AddNewsForm.News_From.value		= AddNewsForm.List_News_From.item(AddNewsForm.List_News_From.selectedIndex).value;
								     break;
								case 3:
										 AddNewsForm.News_KeyWord.value	= AddNewsForm.News_KeyWord.value + "|" + AddNewsForm.List_News_KeyWord.item(AddNewsForm.List_News_KeyWord.selectedIndex).value;
								     break;
								default: 
											//Do Nothing
							}

						} 

						
						
						function HiddenLayer(obj)
						{
							obj.style.visibility="hidden";
							AddNewsForm.List_News_Author.style.visibility	= "hidden";
							AddNewsForm.List_News_From.style.visibility		= "hidden";
							AddNewsForm.List_News_KeyWord.style.visibility= "hidden";
						}
					</script>
					<td width="60" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="ShowOrHideTr('TrForNewsArea')" title="������ʾλ��" bgcolor="#C0C0C0">
					��ʾλ��</td>
					<td width="60" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="ShowOrHideTr('TrForEToolBar1')" title="��չ������һ" bgcolor="#C0C0C0">
					��չ��һ</td>
					<td width="60" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="ShowOrHideTr('TrForEToolBar2')" title="��չ��������" bgcolor="#C0C0C0">
					��չ����</td>
					<td width="*">&nbsp; ��</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td>
			<div id="Layer_News" style="position:absolute; z-index:1; visibility=hidden" onclick="HiddenLayer(this);">
				<table width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td>
							��&nbsp;&nbsp;&nbsp; ��&nbsp;<select size="30" name="List_News_Author" class="InputBox" style="width=125;visibility=hidden" onchange="AddText(1);">
							<%
								Sql = "Select [Author_Name] From [Monolith_News_Author] Order By [Author_Order] Asc"
								Rs.Open Sql, Conn ,1, 1
								Do While Not(Rs.Bof Or Rs.Eof)
							%>
							<option value="<%=Rs(0)%>"><%=Rs(0)%></option>
							<%
									Rs.MoveNext
								Loop
								Rs.Close
						%>

							</select>&nbsp; ��&nbsp;&nbsp; ��&nbsp;<select size="30" name="List_News_From" class="InputBox" style="width=125;visibility=hidden" onchange="AddText(2);">
							<%
								Sql = "Select [From_Name] From [Monolith_News_From] Order By [From_Order] Asc"
								Rs.Open Sql, Conn ,1, 1
								Do While Not(Rs.Bof Or Rs.Eof)
							%>
							<option value="<%=Rs(0)%>"><%=Rs(0)%></option>
							<%
									Rs.MoveNext
								Loop
								Rs.Close
						%>							</select>&nbsp; �� �� ��&nbsp;<select size="30" name="List_News_KeyWord" class="InputBox" style="width=124; visibility=hidden" onchange="AddText(3);">
							<%
								Sql = "Select [KeyWord_Name] From [Monolith_News_KeyWord] Order By [KeyWord_Order] Asc"
								Rs.Open Sql, Conn ,1, 1
								Do While Not(Rs.Bof Or Rs.Eof)
							%>
							<option value="<%=Rs(0)%>"><%=Rs(0)%></option>
							<%
									Rs.MoveNext
								Loop
			Call CloseRs
			Call CloseConn							%>

							</select>
						��</td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
		<tr id="TrForText" style="display:none;">
			<td width="100%" height="21">��������:<textarea rows="4" name="News_Memo" cols="55" class="InputBox" style="width=580" onmouseover="select();"><%=MonolithDecode(News_Memo)%></textarea></td>
		</tr>
		<tr id="TrForPic" style="display:none;">
			<td width="100%">
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
				<tr>
					<td width="260" height="22">ͼ Ƭ һ:<input type="text" name="News_Pic1" size="30" class="InputBox" style="width=200" onmouseover="select();" value="<%=News_Pic1%>"></td>
					<td width="27" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="Insertpic1()">
					����</td>
					<td width="27" class="Menu" align="center" onmouseover="HL_Menu(this,0); View(News_Pic1.value);" onmouseout="HL_Menu(this,1); HiddenLayer();" onclick="HiddenLayer();">
					�쿴</td>
					<td width="20">��</td>
					<td width="210" height="22">ͼ Ƭ ��:<input type="text" name="News_Pic2" size="30" class="InputBox" style="width=150" onmouseover="select();" value="<%=News_Pic2%>"> 
					��</td>
					<td width="27" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="Insertpic2()">
					����</td>
					<td width="27" class="Menu" align="center" onmouseover="HL_Menu(this,0); View(News_Pic2.value);" onmouseout="HL_Menu(this,1); HiddenLayer();" onclick=" HiddenLayer();">
					�쿴</td>
					<td width="*">&nbsp; ��</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td height="21">
			<table height="100%" width="100%" border="0" cellspacing="0" cellpadding="0" background="image/Back.gif">
				<tr>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<!-- ���롢�ϴ�ͼƬ���������ӡ�ȡ���������� -->
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="pic()" title="��������ͼƬ��֧�ָ�ʽΪ��gif��jpg��png��bmp">
					<img src="image/img.gif" border="0"></td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('createLink')" title="��������">
					<img src="image/url.gif" border="0"></td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('unLink')" title="ȡ����������">
					<img src="image/nourl.gif" border="0"></td>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<!-- ���󡢾��С����� -->
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('Justifyleft', '')" title="����">
					<img src="image/Aleft.gif" border="0"> </td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('JustifyCenter', '')" title="����">
					<img src="image/Acenter.gif" border="0"> </td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('JustifyRight', '')" title="����">
					<img src="image/Aright.gif" border="0"> </td>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<!-- ���塢б�塢�»��� -->
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('bold', '')" title="����">
					<img src="image/bold.gif" border="0"> </td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('italic', '')" title="б��">
					<img src="image/italic.gif" border="0"> </td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('underline', '')" title="�»���">
					<img src="image/underline.gif" border="0"> </td>
					<td width="8" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<!-- ȫѡ�� ɾ���ļ���ʽ -->
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('RemoveFormat')" title="ɾ�����ָ�ʽ">
					<img src="image/clear.gif" border="0"></td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('selectall')" title="ȫ��ѡ��">
					<img src="image/selectall.gif" border="0"></td>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<!-- �кš���� -->
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('InsertUnorderedList', '')" title="��Ŀ����">
					<img src="image/list.gif" border="0"></td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('insertorderedlist', '')" title="���">
					<img src="image/num.gif" border="0"></td>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FrameHeight('Add')" title="���ű༭��߶�����50">
					<img src="image/heightAdd.gif" border="0"></td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FrameHeight('Sub')" title="���ű༭��߶ȼ���50">
					<img src="image/heightSub.gif" border="0"></td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="javascript:location.reload()" title="ˢ�¸�ҳ">
					<img src="image/Refresh.gif" border="0"></td>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="Page();" title="�����ҳ��">
					<img src="image/Page.gif" border="0"></td>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<td width="220" class="Menu" align="center">HTML:<input type="text" name="News_FileName" size="24" class="InputBox" value="<%
						if News_Filename <> "" then
							Response.write News_FileName
						else
							Response.write MakeRndNum() & ".htm"
						end if
					%>" onmouseover="select();"></td>
					<td width="*">&nbsp; </td>
				</tr>
			</table>
			</td>
		</tr>
		<tr id="TrForEToolBar1" style="display:none;">
			<td height="22">
			<table height="100%" width="100%" border="0" cellspacing="0" cellpadding="0" background="image/Back.gif">
				<tr>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<!-- ������ �ָ� -->
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('undo')" title="����">
					<img src="image/undo.gif" border="0"> </td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('redo')" title="�ָ�">
					<img src="image/redo.gif" border="0"> </td>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('cut')" title="����">
					<img src="image/cut.gif" border="0"></td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('copy')" title="����">
					<img src="image/copy.gif" border="0"></td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('paste')" title="ճ��">
					<img src="image/paste.gif" border="0"></td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('DELETE')" title="ɾ��">
					<img src="image/del.gif" border="0"></td>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('Outdent', '')" title="����">
					<img src="image/outdent.gif" border="0"></td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('indent', '')" title="����">
					<img src="image/indent.gif" border="0"></td>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('InsertHorizontalRule', '')" title="��ͨˮƽ��">
					<img src="image/line.gif" border="0"></td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('superscript')" title="�ϱ�">
					<img src="image/sup.gif" border="0"></td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="FormatText('subscript')" title="�±�">
					<img src="image/sub.gif" border="0"></td>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="fortable()" title="������">
					<img src="image/table.gif" border="0"></td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="Allmedia()" title="�����ý���ļ�
 1��FLASH
 2����Ƶ�ļ���֧�ָ�ʽΪ��avi��wmv��asf
 3��RealPlay�ļ���֧�ָ�ʽΪ��rm��ra��ram"><img src="image/Media.gif" border="0"></td>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<td width="21" class="Menu" align="center" onmouseover="HL_Menu(this,0)" onmouseout="HL_Menu(this,1)" onclick="help()" title="ʹ�ð���">
					<img src="image/help.gif" border="0"></td>
					<td width="6" class="Menu" align="center">
					<img src="image/Split.gif" border="0"> </td>
					<td width="200" class="Menu" align="center">��������:<input type="text" name="News_Date" size="20" class="InputBox" value="<%
						if News_Date <> "" then
							Response.write News_Date
						else
							Response.write Date()
						end if
					%>" onclick="getDateString(this,oCalendarChs)"></td>
					<td width="*">&nbsp; </td>
				</tr>
			</table>
			</td>
		</tr>
		<tr id="TrForEToolBar2" style="display:none;">
			<td height="22">
			<table height="100%" width="100%" border="0" cellspacing="0" cellpadding="0" background="image/Back.gif">
				<tr>
					<td>
					<select name="selectFont" onchange="FormatText('fontname', selectFont.options[selectFont.selectedIndex].value);document.AddNewsForm.selectFont.options[0].selected = true;" style="font-family: ����; font-size: 9pt" onmouseover="window.status='ѡ��ѡ�����ֵ����塣';return true;" onmouseout="window.status='';return true;">
					<option selected>ѡ������</option>
					<option value="����">����</option>
					<option value="����_GB2312">����</option>
					<option value="������">������</option>
					<option value="����">����</option>
					<option value="����">����</option>
					<option value="��Բ">��Բ</option>
					<option value="Andale Mono">Andale Mono</option>
					<option value="Arial">Arial</option>
					<option value="Arial Black">Arial Black</option>
					<option value="Book Antiqua">Book Antiqua</option>
					<option value="Century Gothic">Century Gothic</option>
					<option value="Comic Sans MS">Comic Sans MS</option>
					<option value="Courier New">Courier New</option>
					<option value="Georgia">Georgia</option>
					<option value="Impact">Impact</option>
					<option value="Tahoma">Tahoma</option>
					<option value="Times New Roman">Times New Roman</option>
					<option value="Trebuchet MS">Trebuchet MS</option>
					<option value="Script MT Bold">Script MT Bold</option>
					<option value="Stencil">Stencil</option>
					<option value="Verdana">Verdana</option>
					<option value="Lucida Console">Lucida Console</option>
					</select>
					<select name="selectColour" onchange="FormatText('ForeColor',selectColour.options[selectColour.selectedIndex].value);document.AddNewsForm.selectColour.options[0].selected = true;" onmouseover="window.status='ѡ��ѡ�����ֵ���ɫ��';return true;" onmouseout="window.status='';return true;">
					<option value="0" selected>ѡ��������ɫ</option>
					<option style="background-color:#FFFFFF;color: #FFFFFF" value="#FFFFFF">
					#FFFFFF</option>
					<option style="background-color:#008000;color: #008000" value="#008000">
					#008000</option>
					<option style="background-color:#800000;color: #800000" value="#800000">
					#800000</option>
					<option style="background-color:#808000;color: #808000" value="#808000">
					#808000</option>
					<option style="background-color:#000080;color: #000080" value="#000080">
					#000080</option>
					<option style="background-color:#800080;color: #800080" value="#800080">
					#800080</option>
					<option style="background-color:#808080;color: #808080" value="#808080">
					#808080</option>
					<option style="background-color:#FFFF00;color: #FFFF00" value="#FFFF00">
					#FFFF00</option>
					<option style="background-color:#00FF00;color: #00FF00" value="#00FF00">
					#00FF00</option>
					<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">
					#00FFFF</option>
					<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">
					#FF00FF</option>
					<option style="background-color:#C0C0C0;color: #C0C0C0" value="#C0C0C0">
					#C0C0C0</option>
					<option style="background-color:#FF0000;color: #FF0000" value="#FF0000">
					#FF0000</option>
					<option style="background-color:#0000FF;color: #0000FF" value="#0000FF">
					#0000FF</option>
					<option style="background-color:#008080;color: #008080" value="#008080">
					#008080</option>
					</select>
					<select name="selectbgColour" onchange="FormatText('BackColor',selectbgColour.options[selectbgColour.selectedIndex].value);document.AddNewsForm.selectbgColour.options[0].selected = true;" onmouseover="window.status='ѡ��ѡ�����ֵı�����ɫ��';return true;" onmouseout="window.status='';return true;">
					<option value="0" selected>ѡ�����ֱ�����ɫ</option>
					<option style="background-color:#FFFFFF;color: #FFFFFF" value="#FFFFFF">
					#FFFFFF</option>
					<option style="background-color:#008000;color: #008000" value="#008000">
					#008000</option>
					<option style="background-color:#800000;color: #800000" value="#800000">
					#800000</option>
					<option style="background-color:#808000;color: #808000" value="#808000">
					#808000</option>
					<option style="background-color:#000080;color: #000080" value="#000080">
					#000080</option>
					<option style="background-color:#800080;color: #800080" value="#800080">
					#800080</option>
					<option style="background-color:#000000;color: #000000" value="#000000">
					#000000</option>
					<option style="background-color:#808080;color: #808080" value="#808080">
					#808080</option>
					<option style="background-color:#0000FF;color: #0000FF" value="#0000FF">
					#0000FF</option>
					<option style="background-color:#FFFF00;color: #FFFF00" value="#FFFF00">
					#FFFF00</option>
					<option style="background-color:#00FF00;color: #00FF00" value="#00FF00">
					#00FF00</option>
					<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">
					#00FFFF</option>
					<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">
					#FF00FF</option>
					<option style="background-color:#C0C0C0;color: #C0C0C0" value="#C0C0C0">
					#C0C0C0</option>
					<option style="background-color:#FF0000;color: #FF0000" value="#FF0000">
					#FF0000</option>
					<option style="background-color:#0000FF;color: #0000FF" value="#0000FF">
					#0000FF</option>
					<option style="background-color:#008080;color: #008080" value="#008080">
					#008080</option>
					</select>
					<select language="javascript" id="FontSize" title="�ֺŴ�С" onchange="FormatText('fontsize',this[this.selectedIndex].value);" name="select" onmouseover="window.status='ѡ��ѡ�����ֵ��ֺŴ�С��';return true;" onmouseout="window.status='';return true;">
					<option class="heading" selected>�ֺ�</option>
					<option value="7">һ��</option>
					<option value="6">����</option>
					<option value="5">����</option>
					<option value="4">�ĺ�</option>
					<option value="3">���</option>
					<option value="2">����</option>
					<option value="1">�ߺ�</option>
					</select>
					<input onclick="setMode(this.checked);" type="checkbox" name="viewhtml" value="ON" id="ViewHtmlSource"><label for="ViewHtmlSource">�鿴HTMLԴ����</label>
					<input onclick="setMode(this.checked);" type="hidden" name="viewhtml" value="ON">
					</td>
				</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100%" height="490" id="TrForBody">
			<script language="javascript">
document.write ('<iframe src="News_Content.asp?News_ID=<%=News_ID%>" id="Monolith_News_Body" width="100%" height="100%" marginwidth="1" marginheight="1"></iframe>')
frames.Monolith_News_Body.document.designMode = "On";
    </script>
			�� 
			<div style="border-style:solid; border-width:1px; position: absolute; width: 90%; height: 100px; z-index: 1; left: 18px; top: 200px; visibility:hidden" id="LayerPicToRemote">
				<div align="center">
					<table border="1" width="100%" id="table1" cellspacing="0" cellpadding="0" height="100%" bordercolor="#CCCCCC">
						<tr>
							<td bgcolor="#C0C0C0">
							<p align="center">�ļ�Զ�̱����� ��<font color="#FF0000">������в���</font>�����Ժ�...</p>
							</td>
						</tr>
					</table>
				</div>
			</div>
			</td>
		</tr>
		<tr>
			<td width="100%">
			<input type="hidden" name="News_ID" value="<%=News_ID%>">
			<input type="hidden" name="News_Content">
			<input type="button" value=" �� �� �� �� (Alt + A) " name="B1" class="InputBox" onclick="AddNews();" accesskey="A">
			<input type="button" value=" Զ �� ͼ Ƭ �� Ϊ �� �� (Alt + S) " name="B3" onclick="SaveToRemote();" class="InputBox" accesskey="S"><iframe name="AAA" width="0" height="0"></iframe>
			</td>
		</tr>
	</table>
</form>

</html>
</head>

