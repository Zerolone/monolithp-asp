<%
	Dim Template_Parent_ID
	Template_Parent_ID = "151"

	Dim MyBBS_Template

	Sql = "Select [Template_Content] From [Monolith_Template] Where [Template_Parent_ID]=" & Template_Parent_ID & " Order By [Template_Level]"
	Rs.Open Sql, Conn, 1, 1
	If Rs.Bof Or Rs.Eof then
		Response.write "模版配置错误，请等待管理更正。 你可以<a href=""#"">点击这里</a>进入简化版本。"
		Response.end
	else

		'Css
		Dim MyBBS_Template_Css, MyBBS_Template_Head, MyBBS_Template_Foot
		MyBBS_Template									= Rs(0)
		MyBBS_Template 									= Split(MyBBS_Template, "<!--Split-->")
		MyBBS_Template_Css							= MyBBS_Template(0)
		MyBBS_Template_Head							= MyBBS_Template(1)
		MyBBS_Template_Foot							= MyBBS_Template(2)

		Rs.MoveNext
		
		'01主页面
		Dim MyBBS_Template_01_Head, MyBBS_Template_01_Body, MyBBS_Template_01_Foot, MyBBS_Template_01_Stats
		MyBBS_Template									= Rs(0)
		MyBBS_Template 									= Split(MyBBS_Template, "<!--Split-->")
		MyBBS_Template_01_Head					= MyBBS_Template(0)
		MyBBS_Template_01_Body					= MyBBS_Template(1)
		MyBBS_Template_01_Foot					= MyBBS_Template(2)		
		MyBBS_Template_01_Stats					= MyBBS_Template(3)
		
		Rs.MoveNext
		
		'02主题列表
		Dim MyBBS_Template_02_Head, MyBBS_Template_02_All_Head, MyBBS_Template_02_Top_Head, MyBBS_Template_02_Top_Body, MyBBS_Template_02_Board_Head, MyBBS_Template_02_Board_Body, MyBBS_Template_02_Foot
		MyBBS_Template									= Rs(0)
		MyBBS_Template 									= Split(MyBBS_Template, "<!--Split-->")
		MyBBS_Template_02_Head					= MyBBS_Template(0)
		MyBBS_Template_02_All_Head			= MyBBS_Template(1)
		MyBBS_Template_02_Top_Head			= MyBBS_Template(2)
		MyBBS_Template_02_Top_Body			= MyBBS_Template(3)
		MyBBS_Template_02_Board_Head		= MyBBS_Template(4)
		MyBBS_Template_02_Board_Body		= MyBBS_Template(5)
		MyBBS_Template_02_Foot					= MyBBS_Template(6) 

		Rs.MoveNext
		
		'03主题察看
		Dim MyBBS_Template_03_Head, MyBBS_Template_03_Body, MyBBS_Template_03_Foot
		MyBBS_Template									= Rs(0)
		MyBBS_Template 									= Split(MyBBS_Template, "<!--Split-->")
		MyBBS_Template_03_Head					= MyBBS_Template(0)
		MyBBS_Template_03_Body					= MyBBS_Template(1)
		MyBBS_Template_03_Foot					= MyBBS_Template(2)
		
		Rs.MoveNext
		
		'04发表主题
		Dim MyBBS_Template_04_Body
		MyBBS_Template									= Rs(0)
		MyBBS_Template 									= Split(MyBBS_Template, "<!--Split-->")
		MyBBS_Template_04_Body					= MyBBS_Template(0)
		
		Rs.MoveNext
		
		'05最新主题列表
		Dim MyBBS_Template_05_Head, MyBBS_Template_05_Body, MyBBS_Template_05_Foot
		MyBBS_Template									= Rs(0)
		MyBBS_Template 									= Split(MyBBS_Template, "<!--Split-->")
		MyBBS_Template_05_Head					= MyBBS_Template(0)
		MyBBS_Template_05_Body					= MyBBS_Template(1)
		MyBBS_Template_05_Foot					= MyBBS_Template(2)

		Rs.MoveNext

		'06回复主题
		Dim MyBBS_Template_06_Head, MyBBS_Template_06_Body, MyBBS_Template_06_Foot
		MyBBS_Template									= Rs(0)
		MyBBS_Template 									= Split(MyBBS_Template, "<!--Split-->")
		MyBBS_Template_06_Head					= MyBBS_Template(0)
		MyBBS_Template_06_Body					= MyBBS_Template(1)
		MyBBS_Template_06_Foot					= MyBBS_Template(2)


	end if
	Rs.Close
	
  'Sql执行次数+1
  SqlQueryNum	= SqlQueryNum + 1
%>