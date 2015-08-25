<!-- #include virtual ="/Inc/Monolith_Conn.asp" -->
<!-- #include virtual ="/Inc/Monolith_Function.asp" -->
<!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Info_ID
	Info_ID					= Request("Info_ID")

	Dim Info_Title,		Info_Title2,	Info_Class_Arr,	Info_Class,		Info_Class_Name,	Info_Author,	Info_From,	Info_KeyWord, 	Info_Memo
	Dim	Info_Pic1,		Info_Pic2, 		Info_FileName,	Info_Content,	Info_Date,		Info_Template,	Info_Area,	Info_Area_Arr,	Info_Area_i,	Info_Order	

	Info_Title			= Request("Info_Title")
	Info_Title2			= Request("Info_Title2")
	Info_Class_Arr	= Split(Request("Info_Class_Arr"), "|")
	Info_Class			= Info_Class_Arr(0)
	Info_Class_Name	= Info_Class_Arr(1)
	
	Info_Author			= Request("Info_Author")
	Info_From				= Request("Info_From")
	Info_KeyWord		= Request("Info_KeyWord")
	Info_Memo				= Request("Info_Memo")
	Info_Pic1				= Request("Info_Pic1")
	Info_Pic2				= Request("Info_Pic2")
	Info_FileName		= Request("Info_FileName")
	Info_Content		= Request("Info_Content")
	Info_Date				= Request("Info_Date")
	Info_Template		= Request("Info_Template")
	Info_Area_Arr		= Request("Info_Area")
	Info_Order			= Request("Info_Order")

	Info_Title			= MonolithEncode(Info_Title)
	Info_Title2			= MonolithEncode(Info_Title2)
	Info_Author			= MonolithEncode(Info_Author)
	Info_From				= MonolithEncode(Info_From)
	Info_KeyWord		= MonolithEncode(Info_KeyWord)
	Info_Memo				= MonolithEncode(Info_Memo)
	
	Info_Area				= 0
	Info_Area_Arr = Split(Info_Area_Arr, ",")
	For Info_Area_i=LBound(Info_Area_Arr) to UBound(Info_Area_Arr)
		Info_Area = Info_Area + CLng(Info_Area_Arr(Info_Area_i))
	Next
	
	if not IsNumeric(Info_Order) then Info_Order = 999
	
	if Info_ID <> "" then
		'------------------0-------------1--------------2-------------3------------------4--------------5------------6---------------7------------8------------9------------10---------------11--------------12-----------13---------------14-----------15------------16
		Sql = "Select [Info_Title], [Info_Title2], [Info_Class], [Info_Class_Name], [Info_Author], [Info_From], [Info_KeyWord], [Info_Memo], [Info_Pic1], [Info_Pic2], [Info_FileName], [Info_Content], [Info_Date], [Info_Template], [Info_Area], [Info_Order], [Info_Apply] From [Monolith_Info] Where [Info_ID]=" & Info_ID
		Rs.Open Sql, Conn, 1, 3
		if Not(Rs.Bof Or Rs.Eof) then
			Rs(0)				= Info_Title
			Rs(1)				= Info_Title2
			Rs(2)				= Info_Class
			Rs(3)				= Info_Class_Name
			Rs(4)				= Info_Author
			Rs(5)				= Info_From
			Rs(6)				= Info_KeyWord
			Rs(7)				= Info_Memo
			Rs(8)				= Info_Pic1
			Rs(9)				= Info_Pic2
			Rs(10)			= Info_FileName
			Rs(11)			= Info_Content
			Rs(12)			= Info_Date
			Rs(13)			= Info_Template
			Rs(14)			= Info_Area
			Rs(15)			= Info_Order
			Rs(16)			= 0
		end if
		Rs.Update
		Call CloseRs
		Call CloseConn
		Response.write "<script language=""javascript"">"
		Response.write "  alert(""修改成功， 正在挑转到编辑页面， 请确认无误后进行审核！"");"
		Response.write "  window.open('Add.asp?Info_ID=" & Info_ID & "','_self')"
		Response.write "</script>"

	else
		'------------------0-------------1--------------2-------------3------------------4--------------5------------6---------------7------------8------------9------------10---------------11--------------12-----------13---------------14-----------15
		Sql = "Select [Info_Title], [Info_Title2], [Info_Class], [Info_Class_Name], [Info_Author], [Info_From], [Info_KeyWord], [Info_Memo], [Info_Pic1], [Info_Pic2], [Info_FileName], [Info_Content], [Info_Date], [Info_Template], [Info_Area], [Info_Order] From [Monolith_Info]"
		Rs.Open Sql, Conn, 1, 3
		Rs.AddNew
		Rs(0)				= Info_Title
		Rs(1)				= Info_Title2
		Rs(2)				= Info_Class
		Rs(3)				= Info_Class_Name
		Rs(4)				= Info_Author
		Rs(5)				= Info_From
		Rs(6)				= Info_KeyWord
		Rs(7)				= Info_Memo
		Rs(8)				= Info_Pic1
		Rs(9)				= Info_Pic2
		Rs(10)			= Info_FileName
		Rs(11)			= Info_Content
		Rs(12)			= Info_Date
		Rs(13)			= Info_Template
		Rs(14)			= Info_Area
		Rs(15)			= Info_Order
		Rs.Update
		Call CloseRs
		Call CloseConn
		Response.write "<script language=""javascript"">"
		Response.write "  alert(""添加成功, 请刷新页面察看新添加的数据。"");"
		Response.write "  window.top.Frame_Right.win.removewin(window.top.Frame_Right.win.currentwin);"
		Response.write "</script>"
	end if
%>