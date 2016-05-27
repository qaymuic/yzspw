<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Option Explicit
Response.Buffer = True
Dim Startime
Dim Dvbbs,template
Dim SqlNowString,Conn
'定义数据库类别，1为SQL数据库，0为Access数据库
Const IsSqlDataBase = 0
'论坛缓存名称，如果一个站点有多个论坛请更改成不同名称
Const Forum_CacheName = "aspsky"
'定义运行模式，测试的时候设置1，正常运行的时候设置为0,不输出错误信息有利于安全，
Const IsDeBug = 1
Startime = Timer()
Dim MyBoardOnline
Dim Db
Set Dvbbs = New Cls_Forum
Set MyBoardOnline=new Cls_UserOnlne 
Set template = New cls_templates
If IsSqlDataBase = 1 Then
	SqlNowString = "GetDate()"
Else
	SqlNowString = "Now()"
End If
Sub ConnectionDatabase
	Dim ConnStr
	If IsSqlDataBase = 1 Then
		'sql数据库连接参数：数据库名、用户密码、用户名、连接名（本地用local，外地用IP）
		Dim SqlDatabaseName,SqlPassword,SqlUsername,SqlLocalName
		SqlDatabaseName = "dvbbs7"
		SqlPassword = ""
		SqlUsername = "dvbbs"
		SqlLocalName = "(local)"
		ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlLocalName & ";"
	Else
		'免费用户第一次使用请修改本处数据库地址并相应修改data目录中数据库名称，如将dvbbs6.mdb修改为dvbbs6.asp
		Db = "data/dvbbs7.mdb"
		ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)
	End If
	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.open ConnStr
	If Err Then
		err.Clear
		Set Conn = Nothing
		Response.Write "数据库连接出错，请检查连接字串。"'注释，需要把这几个字翻译成英文。
		Response.End
	End If
End Sub
%>
