<%@language=vbscript codepage=936 %>
<%
option explicit
response.buffer=true	
dim conn
dim connstr
dim db
db="database/yzspwdb.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
conn.Open connstr
 If Err Then
	err.Clear
	Set Conn = Nothing
	Response.Write "���ݿ����ӳ������������ִ���"'ע�ͣ���Ҫ���⼸���ַ����Ӣ�ġ�
	Response.End
 End If

sub CloseConn()
	conn.close
	set conn=nothing
end sub
%>

