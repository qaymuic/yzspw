<!--#Include File="Dv_ClsMain.asp"-->
<%
'是否商业版，非官方SQL版本请在此设置为0以及在Conn中设置论坛为SQL数据库，否则显示不正常
Const IsBuss=1
Dvbbs.GetForum_Setting
Dvbbs.CheckUserLogin
%>