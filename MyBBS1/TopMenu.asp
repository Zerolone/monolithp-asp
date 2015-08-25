<style>
<!--
.tableborder { border: 1px solid #345487; background-color: #FFF; padding: 0px; margin: 0px; width: 100% }
.maintitle { vertical-align: middle; font-weight: bold; color: #FFF; letter-spacing: 1px; background-image: url('images/tile_back.gif'); padding-left:5px; padding-right:0px; padding-top:8px; padding-bottom:8px }
.titlemedium { font-weight: bold; color: #3A4F6C; padding: 7px; margin: 0px; background-image: url('images/tile_sub.gif') }
td { font-family: Tahoma,Verdana, Tahoma, Arial,Helvetica, sans-serif; font-size: 12px; color: #000; word-break: break-all }
.row4 { background-color: #E4EAF2 }
.desc { font-size: 12px; color: #434951 }
.row2 { background-color: #DFE6EF }
.darkrow2 { background-color: #BCD0ED; color: #3A4F6C }
-->
</style>
<table border="1" width="100%">
  <tr> 
    <td> <%
dim UserName , Password
if Request.Cookies.count<>0 then
'  Username=replace(request.Cookies(MyBBSCookiesName)("Username"),"'","")
'  Password=replace(request.Cookies(MyBBSCookiesName)("Password"),"'","")

  Username=request.Cookies(MyBBSCookiesName)("Username")
  Password=request.Cookies(MyBBSCookiesName)("Password")

  sql= "SELECT * FROM [MyBBS_Users] where [Username] = '" & Username & "' and [password]='" & Password & "'"
  set rs=conn.Execute(sql)
  if not (rs.BOF or rs.eof) then
    TheString	= UserName & "，欢迎回来！<a href='Logout.asp'>点这里退出</a>"
	IsLogin		= 1
  end if
  rs.Close
end if
if IsLogin <> 1 then  TheString = "<a href='Reg.asp'>用户注册</a>&nbsp;&nbsp;<a href='Login.asp'>用户登录</a>"
Response.write TheString
%> </td>
    <td> <p><a href="index.asp">首页</a></p></td>
    <td> <p><a href="http://www.showde.com/zero/index.asp?vt=bycat&cat_id=49" target="_blank">制作进度</a></p></td>
    <td> <p><a href="AboutUs.asp">关于我们</a></p></td>
  </tr>
</table>
<% if Username="Zerolone" then %>
<table width="100%" border="1" bgcolor="#CCCCCC">
  <tr> 
    <td><a href="AdminBoardAdd.asp">添加板块</a></td>
    <td><a href="AdminBoard.asp">所有板块/排序</a></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><a href="AdminAboutUs.asp">修改关于我们</a></td>
  </tr>
</table>
<%end if%>