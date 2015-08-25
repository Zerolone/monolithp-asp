<!--#include file="../Inc/DigiChina_Conn.asp"-->
<!--#include file="../Inc/MD5.asp"--><!-- #include virtual ="/Manage/Check.asp"-->
<%
	Dim Admin_Name, Admin_Password
	Admin_Name			= Replace(Request.Form("Admin_Name"),"'","''")
	Admin_Password	= Replace(Request.Form("Admin_Password"),"'","''")
 	Admin_Password	= MD5(Admin_Password,32)

  Sql = "Select [Admin_Name] From [DigiChina_Admin] where [Admin_Name] ='" & Admin_Name & "' And [Admin_Password] = '" & Admin_Password & "'"
  Rs.Open Sql, Conn, 1, 3
  If Rs.Bof Or Rs.Eof or Session("Admin_Code")<>Admin_Code or not IsNumeric(session("Admin_Code")) then
		CanLogin=false
		Session("Admin_Name")			= ""
    Session("Admin_Password")	= ""
		Session("Admin_Login")		= 0
    Session.Abandon
		Response.Write "&login_sus=0&"
  else
    Session("Admin_Name")			= Admin_Name
		Session("Admin_Password")	= Admin_Password
		Session("Admin_Login")		= 1
		CanLogin=true
		Response.Write"&login_sus=1&"
  end if
	Rs.close
	Set Rs=Nothing
	Call CloseDB
%>