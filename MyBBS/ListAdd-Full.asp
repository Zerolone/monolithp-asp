<!-- #include file ="../Inc/Monolith_Conn.asp" -->
<!-- #include file="TopMenu.asp"-->
<!-- #include file="Skin.asp"--><%
  if Username="" then
%>
<!-- #include file="Logini.asp"--><%
  end if

  Dim Board_ID
   
  Board_ID		= Request("Board_ID")

  dim ThisPage, FormPage
  ThisPage = "ListSave.asp"
  FormPage	= Request.ServerVariables("HTTP_REFERER")
%>
<!--分割线--><img src="images/Blank.gif" border="0" width="1" height="5"><!--分割线--><table width="100%" cellspacing="0" cellpadding="0" style="border: 1px solid #000000">
	<form action="<%=ThisPage%>" method="post" name="ListAddFrm" id="ListAddFrm">
		<input name="BBS_Board_ID" type="hidden" value="<%=Board_ID%>">
		<input name="FormPage" type="hidden" value="<%=FormPage%>">
		<tr>
			<td background="images/tile_back.gif" height="32" colspan="2"><b>
			<font color="#FFFFFF">&nbsp;&nbsp; 在 {Board_Name} 发表新的主题</font></b>
			</td>
		</tr>
		<tr>
			<td class="pformstrip" colspan="2">主题设置</td>
		</tr>
		<tr>
			<td class="pformleft">主题名称</td>
			<td class="pformright">
			<input class="forminput" tabindex="1" maxlength="50" size="40" name="BBS_Title"></td>
		</tr>
		<tr>
			<td class="pformleft">主题说明</td>
			<td class="pformright">
			<input class="forminput" tabindex="2" maxlength="40" size="40" name="TopicDesc"></td>
		</tr>
		<tr>
			<td class="pformstrip" colspan="2">代码编辑按钮</td>
		</tr>
		<tr>
			<td class="pformleft">&nbsp;　</td>
			<td class="pformright" valign="top">按钮<br>
			放这里 </td>
		</tr>
		<tr>
			<td class="pformstrip" colspan="2">主题内容</td>
		</tr>
		<tr>
			<td class="pformleft" align="middle">
			<table style="WIDTH: auto" cellspacing="0" id="table2">
				<tr>
					<td align="middle" colspan="3"><b>Clickable Smilies</b></td>
				</tr>
				<tr align="middle">
					<td><a href="javascript:emoticon('<_<')">
					<img alt="smilie" src="style_emoticons/default/dry.gif" border="0"></a>
					</td>
					<td><a href="javascript:emoticon('-_-')">
					<img alt="smilie" src="style_emoticons/default/sleep.gif" border="0"></a>
					</td>
					<td><a href="javascript:emoticon(':(')">
					<img alt="smilie" src="style_emoticons/default/sad.gif" border="0"></a>
					</td>
				</tr>
				<tr align="middle">
					<td><a href="javascript:emoticon(':)')">
					<img alt="smilie" src="style_emoticons/default/smile.gif" border="0"></a>
					</td>
					<td><a href="javascript:emoticon(':D')">
					<img alt="smilie" src="style_emoticons/default/biggrin.gif" border="0"></a>
					</td>
					<td><a href="javascript:emoticon(':P')">
					<img alt="smilie" src="style_emoticons/default/tongue.gif" border="0"></a>
					</td>
				</tr>
				<tr align="middle">
					<td><a href="javascript:emoticon(':angry:')">
					<img alt="smilie" src="style_emoticons/default/mad.gif" border="0"></a>
					</td>
					<td><a href="javascript:emoticon(':blink:')">
					<img alt="smilie" src="style_emoticons/default/blink.gif" border="0"></a>
					</td>
					<td><a href="javascript:emoticon(':huh:')">
					<img alt="smilie" src="style_emoticons/default/huh.gif" border="0"></a>
					</td>
				</tr>
				<tr align="middle">
					<td><a href="javascript:emoticon(':lol:')">
					<img alt="smilie" src="style_emoticons/default/laugh.gif" border="0"></a>
					</td>
					<td><a href="javascript:emoticon(':mellow:')">
					<img alt="smilie" src="style_emoticons/default/mellow.gif" border="0"></a>
					</td>
					<td><a href="javascript:emoticon(':o')">
					<img alt="smilie" src="style_emoticons/default/ohmy.gif" border="0"></a>
					</td>
				</tr>
				<tr align="middle">
					<td><a href="javascript:emoticon(':ph34r:')">
					<img alt="smilie" src="style_emoticons/default/ph34r.gif" border="0"></a>
					</td>
					<td><a href="javascript:emoticon(':rolleyes:')">
					<img alt="smilie" src="style_emoticons/default/rolleyes.gif" border="0"></a>
					</td>
					<td><a href="javascript:emoticon(':unsure:')">
					<img alt="smilie" src="style_emoticons/default/unsure.gif" border="0"></a>
					</td>
				</tr>
				<tr align="middle">
					<td><a href="javascript:emoticon(':wacko:')">
					<img alt="smilie" src="style_emoticons/default/wacko.gif" border="0"></a>
					</td>
					<td><a href="javascript:emoticon(':wub:')">
					<img alt="smilie" src="style_emoticons/default/wub.gif" border="0"></a>
					</td>
					<td><a href="javascript:emoticon(';)')">
					<img alt="smilie" src="style_emoticons/default/wink.gif" border="0"></a>
					</td>
				</tr>
				<tr align="middle">
					<td><a href="javascript:emoticon('B)')">
					<img alt="smilie" src="style_emoticons/default/cool.gif" border="0"></a>
					</td>
					<td><a href="javascript:emoticon('^_^')">
					<img alt="smilie" src="style_emoticons/default/happy.gif" border="0"></a>
					</td>
					<td>　</td>
				</tr>
				<tr>
					<td align="middle" colspan="3"><b>
					<a href="javascript:emo_pop()">全部显示</a></b></td>
				</tr>
			</table>
			</td>
			<td class="pformright" valign="top">
			<textarea class="textarea" tabindex="5" name="BBS_Content" rows="20" cols="80"></textarea></td>
		</tr>
		<tr>
			<td class="pformleft"><b>主题选项</b></td>
			<td class="pformright"><br>
			<input class="checkbox" type="checkbox" checked value="yes" name="enableemo">
			<b>使用 emoticons?</b> <br>
			<input class="checkbox" type="checkbox" checked value="yes" name="enablesig">
			<b>使用 签名</b> <br>
			<input class="checkbox" type="checkbox" value="1" name="enabletrack">
			<b>使用 电子邮件回复通知么? </b></td>
		</tr>
		<tr>
			<td class="pformleft"><b>Post Icons</b></td>
			<td class="pformright">
			<input class="radiobutton" type="radio" value="V1" name="iconid1" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon1.gif" align="middle">&nbsp;&nbsp;
			<input class="radiobutton" type="radio" value="V1" name="iconid2" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon2.gif" align="middle">&nbsp;&nbsp;
			<input class="radiobutton" type="radio" value="V1" name="iconid3" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon3.gif" align="middle">&nbsp;&nbsp;
			<input class="radiobutton" type="radio" value="V1" name="iconid4" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon4.gif" align="middle">&nbsp;&nbsp;
			<input class="radiobutton" type="radio" value="V1" name="iconid5" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon5.gif" align="middle">&nbsp;&nbsp;
			<input class="radiobutton" type="radio" value="V1" name="iconid6" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon6.gif" align="middle">&nbsp;&nbsp;
			<input class="radiobutton" type="radio" value="V1" name="iconid7" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon7.gif" align="middle"><br>
			<input class="radiobutton" type="radio" value="V1" name="iconid8" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon8.gif" align="middle">&nbsp;&nbsp;
			<input class="radiobutton" type="radio" value="V1" name="iconid9" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon9.gif" align="middle">&nbsp;&nbsp;
			<input class="radiobutton" type="radio" value="V1" name="iconid10" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon10.gif" align="middle">&nbsp;&nbsp;
			<input class="radiobutton" type="radio" value="V1" name="iconid11" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon11.gif" align="middle">&nbsp;&nbsp;
			<input class="radiobutton" type="radio" value="V1" name="iconid12" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon12.gif" align="middle">&nbsp;&nbsp;
			<input class="radiobutton" type="radio" value="V1" name="iconid13" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon13.gif" align="middle">&nbsp;&nbsp;
			<input class="radiobutton" type="radio" value="V1" name="iconid14" checked>&nbsp;
			<img alt="" src="folder_post_icons/icon14.gif" align="middle"><br>
			<input class="radiobutton" type="radio" checked value="V1" name="iconid">&nbsp; 
			[ Use None ] </td>
		</tr>
		<tr>
			<td class="pformstrip" colspan="2">File Attachments</td>
		</tr>
		<tr>
			<td class="pformleft" valign="top">你可以在此篇文章中附带文件<br>
			文件大小限制为:: 1.95mb</td>
			<td class="pformright">
			<input class="forminput" type="file" size="30" name="FILE_UPLOAD">
			<input class="button" onclick="Override=1;" type="submit" value="增加这个附件" name="attachgo"></td>
		</tr>
		<tr>
			<td class="formbuttonrow" colspan="2">
			<p align="center">
			<input class="button" accesskey="s" tabindex="7" type="submit" value="Post New Topic" name="submit0">&nbsp;
			<input class="button" tabindex="8" type="submit" value="Preview Post" name="preview">
			</p>
			</td>
		</tr>
	</form>
</table>
<!-- #include file="CopyRight.asp"-->