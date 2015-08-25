<%
const MaxUser=100

function OnlyLogin
	if session("userid")="" then
		session("strURL")="welcome.htm"
		session("strMsg")="此功能只有系统用户才能使用，请先登录或注册！"
		response.redirect "alert.asp"
		response.End
	end if
end function

function OnlySysOp
	if not isSysOp(session("userid")) then
		session("strURL")="Welcome.htm"
		session("strMsg")="此功能只有系统管理员才能使用！"
		response.redirect "alert.asp"
		response.End
	end if
end function

function isBMaster(bstr)
	dim ret,chstr
	ret=false
	bstr=Lcase(bstr)
	chstr="," & Lcase(session("userid")) & ","
	if instr(1,bstr,chstr)>0 then ret=true
	isBmaster=ret
end function

function isSysOP(str)
		dim ret,chstr
		ret=false
		chstr=Lcase(str)&","
		if instr(1,Application("Sysop"),chstr)>0 then
			ret=true
		end if
		isSysOp=ret
end function

function Ginit()
		dim rsSysinfo1,sysTitle
		if Application("hasUserList")=0 then
			dim i
			dim UserList(100,5)
			for i=1 to maxUser-1
				userList(i,0)=""
				UserList(i,1)=timer()
				UserList(i,2)=timer()
				UserList(i,3)=""
				UserList(i,4)="白痴中"
			next 
			set rsSysinfo1=server.createobject("ADODB.RECORDSET")

			rsSysinfo1.Open "select * from sysinfo",myconn,1
			if isNull(rsSysinfo1("title")) then
				sysTitle="新时代社区"
			else
				sysTitle=rsSysinfo1("title")
			end if
			rsSysinfo1.Close

			Application.Lock
			Application("UserList") = UserList
			Application("hasUserList")=1
			Application("Online")=0
			Application("Title")=sysTitle
			Application.Unlock
		end if
end function

function Alive(str)
	if session("userid")<>"" then
		if session("userListID")>=0 then
			dim UserList
			UserList = Application("UserList")
			UserList(session("UserListID"),2)=timer()
			UserList(session("UserListID"),4)=str
			Application.Lock
			Application("UserList")=UserList
			Application.Unlock
		end if
	else
		session("strURL")="kick.htm"
		session("strMsg")="由于发呆时间过长，你已经被踢出，请重新登录！"
		response.redirect "alert.asp"
		response.End
	end if
end function

function initLogin()
		Ginit
		dim UserList,i,blDone
		dim UserUp
		dim UserStay
		dim rsSysinfo1,sysTitle
		dim ison
		ison=0
		blDone=false
		UserList = Application("UserList")
		for i=1 to maxUser-1
			if not blDone and UserList(i,0)="" then
				UserList(i,0)=session("userid")
				UserList(i,3)=session("nickname")
				UserList(i,1)=timer()
				UserList(i,2)=timer()
				session("userListID")=i
				blDone=true
			end if
			     if UserList(i,0)=session("userid") then
				UserList(i,0)=session("userid")
				UserList(i,3)=session("nickname")
				UserList(i,1)=timer()
				UserList(i,2)=timer()
				session("userListID")=i
				blDone=true
				ison=1
				end if
		next

		Application.Lock
		Application("UserList")=UserList
		if ison=0 then
		Application("Online")=Application("Online")+1
		Application("counter")=Application("counter")+1
		end if
		Application.UnLock
		
		session("ListMethod")=1 	'	  =1  动态方式   =2  时间方式
		session("MenuMethod")=1  	'	  =1  网易分类   =2  列表方式
		session("OnePage")=15

		dim rsSysInfo
		set rsSysInfo=Server.CreateObject("ADODB.RecordSet")
		rsSysInfo.Open "select * from SysInfo",myconn,1
		if rsSysInfo.RecordCount>0 then		
			Application.Lock
			Application("SysOP")=Lcase(rsSysInfo("SysOP"))
			Application.Unlock
		end if
		rsSysInfo.Close
		session.TimeOut=40
end function

function quitLogin()
		dim UserList
		if session("userListID")>0 then
			UserList = Application("UserList")
			UserList(session("userListID"),0)=""
			
			Application.Lock
			Application("UserList")=UserList
			Application("Online")=Application("Online")-1
			Application.Unlock
		end if
		session("userid")=""
		session.Abandon
end function

%>

