<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim News_ID
	News_ID					= Request("News_ID")

	Dim News_Title,		News_Title2,	News_Class, 		News_Author,	News_From,	News_KeyWord, 	News_Memo, News_Pic1, News_Pic2, News_FileName
	Dim News_Content,	News_Date,		News_Template,	News_Area, 		News_Area_Arr,	News_Area_i,	News_Order

	News_Title			= Request("News_Title")
	News_Title2			= Request("News_Title2")
	News_Class			= Request("News_Class")
	News_Author			= Request("News_Author")
	News_From				= Request("News_From")
	News_KeyWord		= Request("News_KeyWord")
	News_Memo				= Request("News_Memo")
	News_Pic1				= Request("News_Pic1")
	News_Pic2				= Request("News_Pic2")
	News_FileName		= Request("News_FileName")
	News_Content		= Request("News_Content")
	News_Date				= Request("News_Date")
	News_Template		= Request("News_Template")
	News_Area_Arr		= Request("News_Area")
	News_Order			= Request("News_Order")

	News_Title			= MonolithEncode(News_Title)
	News_Title2			= MonolithEncode(News_Title2)
	News_Author			= MonolithEncode(News_Author)
	News_From				= MonolithEncode(News_From)
	News_KeyWord		= MonolithEncode(News_KeyWord)
	News_Memo				= MonolithEncode(News_Memo)
	
	News_Area				= 0
	News_Area_Arr = Split(News_Area_Arr, ",")
	For News_Area_i=LBound(News_Area_Arr) to UBound(News_Area_Arr)
		News_Area = News_Area + CLng(News_Area_Arr(News_Area_i))
	Next
	
	if not IsNumeric(News_Order) then News_Order = 999
	
	Response.write "<script language=""javascript"">"
	if News_ID <> "" then
		'------------------0-------------1--------------2-------------3--------------4------------5---------------6------------7------------8------------9----------------10--------------11-----------12---------------13-----------14------------15
		Sql = "Select [News_Title], [News_Title2], [News_Class], [News_Author], [News_From], [News_KeyWord], [News_Memo], [News_Pic1], [News_Pic2], [News_FileName], [News_Content], [News_Date], [News_Template], [News_Area], [News_Order], [News_Apply] From [Monolith_News] Where [News_ID]=" & News_ID
		Rs.Open Sql, Conn, 1, 3
		if Not(Rs.Bof Or Rs.Eof) then
			Rs(0)				= News_Title
			Rs(1)				= News_Title2
			Rs(2)				= News_Class
			Rs(3)				= News_Author
			Rs(4)				= News_From
			Rs(5)				= News_KeyWord
			Rs(6)				= News_Memo
			Rs(7)				= News_Pic1
			Rs(8)				= News_Pic2
			Rs(9)				= News_FileName
			Rs(10)			= News_Content
			Rs(11)			= News_Date
			Rs(12)			= News_Template
			Rs(13)			= News_Area
			Rs(14)			= News_Order
			Rs(15)			= 0
		end if
		Rs.Update
		Call CloseRs
		Response.write "  alert(""修改成功， 正在挑转到编辑页面， 请确认无误后进行审核！"");"
	else
		'------------------0-------------1--------------2-------------3--------------4------------5---------------6------------7------------8------------9----------------10--------------11-----------12---------------13-----------14------------15
		Sql = "Select [News_Title], [News_Title2], [News_Class], [News_Author], [News_From], [News_KeyWord], [News_Memo], [News_Pic1], [News_Pic2], [News_FileName], [News_Content], [News_Date], [News_Template], [News_Area], [News_Order], [News_ID] From [Monolith_News]"
		Rs.Open Sql, Conn, 1, 3
		Rs.AddNew
		Rs(0)				= News_Title
		Rs(1)				= News_Title2
		Rs(2)				= News_Class
		Rs(3)				= News_Author
		Rs(4)				= News_From
		Rs(5)				= News_KeyWord
		Rs(6)				= News_Memo
		Rs(7)				= News_Pic1
		Rs(8)				= News_Pic2
		Rs(9)				= News_FileName
		Rs(10)			= News_Content
		Rs(11)			= News_Date
		Rs(12)			= News_Template
		Rs(13)			= News_Area
		Rs(14)			= News_Order
		Rs.Update
		News_ID			= Rs(15)
		Call CloseRs
		Response.write "  alert(""添加成功， 正在挑转到编辑页面， 请确认无误后进行审核！"");"
	end if
	
	Call CloseConn

	Response.write "  window.open('Add.asp?News_ID=" & News_ID & "','_self')"
	Response.write "</script>"
%>