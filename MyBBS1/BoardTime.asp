<%
  Endtime = Timer()
  TheString = "<hr>ִ��ʱ�䣺" & FormatNumber((Endtime-Startime)*1000,5) & "���롣��ѯ���ݿ�" & SqlQueryNum & "��"
  Response.write TheString
%>