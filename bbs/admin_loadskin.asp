<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%Head()%>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall'){
	e.checked = form.chkall.checked;
	}
   }
  }
//-->
</script>
<table border="0" cellspacing="1" cellpadding="5" align=center width="95%" class="tableBorder">
<tr>
<th colspan="3" align="center" ID="TableTitleLink"><a href=?>��̳ģ�浼������</a> | <a href=?action=load>��̳ģ�浼�빦��</a></th>
</tr>
<tr>
<td class="forumrow">
ע��<br>
����ȷ��ģ�����ݿ�����ȷ��<br>
������ģ�����ݿ����skinĿ¼�£�����д��Skins/Dv_skin.mdb��<br>
����ģ�����ݿ��ڱ��ݵı���ΪDv_Style,�벻Ҫ���ģ�<br>
����ģ�����ݰ�����̳�ãӣ����ã��뼰������̳ͼƬ���ã�
</td>
</tr>
</table><br>
<%
Dim admin_flag
Dim skid,sname,act,mdbname,StyleConn,SucMsg
admin_flag=",21,"
If not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 Then
	Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	dvbbs_error()
Else
	If Request("action")="inputskin" Then
		Call inputskin()
	ElseIf Request("action")="loadskin" Then
		Call loadskin()
	ElseIf Request("action")="load" Then
		Call load()
	ElseIf Request("action")="rename" Then
		Call rename()
	ElseIf Request("action")="savenm" Then
		Call savenm()
	ElseIf Request("action")="CreatMdb" Then
		Call CreateStyleMdb()
	ElseIf Request("action")="DelFields" Then
		Call DelFields()
	Else
		Call MAIN()
	End If
End If
If Errmsg<>"" Then dvbbs_error()
If IsObject(StyleConn) Then
	StyleConn.close
	Set StyleConn=Nothing
End IF
Call Footer()

Sub MAIN()
If Request("action")="loadthis" Then
    sname="����"
    act="loadskin"
    mdbname=Dvbbs.Checkstr(trim(Request.form("skinmdb")))
    If mdbname="" Then
		Errmsg=ErrMsg + "<li>����д����ģ�汣��ı���"
		Exit Sub
	End If
Else
	sname="����"
	act="inputskin"
End If
%>
<table border="0" cellspacing="1" cellpadding="5" align=center width="95%" class="tableBorder">
<tr><th width="100%" colspan="4"><%=sname%>��̳ģ���б�</th></tr>
<tr>
<td width="10%"  align="center" class="forumrow">���</td>
<td width="65%"  align="center" class="forumrow">ģ������</td>
<td width="20%"  align="center" class="forumrow">����</td>
<td width="5%"  align="center" class="forumrow">ѡ��</td>
</tr>
<form action="?action=<%=act%>" method=post name=even>
<%
If act="loadskin" Then
	SkinConnection(mdbname)
	set Rs=StyleConn.Execute("select id,StyleName from Dv_Style order by id ")
Else
	set Rs=Dvbbs.Execute("select id,StyleName from Dv_Style order by id ")
End If
	do while not Rs.eof
%>
<tr>
	<td class="forumrow"><%=Rs("id")%></td>
	<td class="forumrow"><%=Rs("StyleName")%></td>
	<td class="forumrow" align=center>
	<a href="?action=rename&act=<%=act%>&skid=<%=Rs("id")%>&mdbname=<%=mdbname%>" >����</a>
	<%If act<>"loadskin" Then
	Response.Write " | <a href=""admin_template.asp?action=manage&mostyle=�� ��&StyleID="&Rs("id")&""" >�༭</a>"
	End If %>
	</td>
	<td class="forumrow" align=center><input type="checkbox" name="skid" value="<%=Rs("id")%>"></td>
</tr>
<%	Rs.movenext
	loop
	Rs.close:Set Rs=Nothing
%>
<tr>
<td colspan="4" align=center class="forumRowHighlight">
<%=sname%>�����ݿ⣺<input type="text" name="skinmdb" size="30" value="Skins/Dv_skin.mdb">
<input type="submit" name="submit" value="<%=sname%>">
<input type=submit name=Submit value=ɾ��  onclick="{if(confirm('ע�⣺��ɾ����ģ�潫���ָܻ���')){this.document.even.submit();return true;}return false;}">  <input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">ȫѡ</td>
</tr>
</form>
</table>
<% 
End Sub

Sub SkinConnection(mdbname)
On Error Resume Next 
	Set StyleConn = Server.CreateObject("ADODB.Connection")
	StyleConn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
	If Err.Number ="-2147467259"  Then 
		Errmsg=ErrMsg + "<li>"&mdbname&"���ݿⲻ���ڡ�"
		Dvbbs_error()
		Response.end
	End If
End Sub

Sub inputskin()
Dim TempRs
	skid=Dvbbs.checkstr(Request("skid"))
   	mdbname=Dvbbs.Checkstr(Trim(Request.form("skinmdb")))
If skid="" or Isnull(skid) or Not IsNumeric(Replace(Replace(skid,",","")," ","")) Then
	Errmsg=ErrMsg + "<li>����δѡȡҪ������ģ�棬������д���"
	Exit Sub
End If
If mdbname="" Then
	Errmsg=ErrMsg + "<li>������д����ģ�����ݿ���"
	Exit Sub
End If

If Request("submit")="ɾ��" Then
	If instr(","&skid&",",","&Dvbbs.cachedata(17,0)&",") Then 
		Errmsg=ErrMsg + "<BR><li>��ģ����Ĭ��ģ�棬������ɾ����"
		Exit Sub
	End If
	Set Rs=Dvbbs.Execute("select Count(*) From [Dv_Board] Where sid in ("&skid&")")
	If Rs(0)>0 Then
		Set Rs=Nothing 
		Errmsg=ErrMsg + "<BR><li>��ģ�����з���̳��ʹ�ã�����ɾ����"
		Dvbbs_error()
	End If
	Set rs=Nothing 
	Dvbbs.Execute("Delete From [Dv_Style] Where ID in ("&skid&")")
	Dv_suc("�ɹ�ɾ��ģ�塣")
	Dvbbs.DelCahe("Templateslist")
	'ɾ����ģ������ҳ�滺��
	Set Rs=Dvbbs.Execute("Select Top 0 * From [Dv_Style]")
	For i=2 to Rs.Fields.Count-1
		Dvbbs.DelCahe(Rs(i).Name&skid)	
	Next
	Dvbbs.DelCahe("BbsListTop"&skid)
	Set Rs=Nothing
Else
	SkinConnection(mdbname)
	ChkSkinMDB()
	If Errmsg<>"" Then Exit Sub
	set Rs=Dvbbs.Execute("select * from Dv_Style where id in ("&skid&") order by id ")
	If Rs.EOF Or Rs.BOF Then
		Errmsg=ErrMsg + "<BR><li>�޷�ȡ��Դģ������"
		Dvbbs_error()
		Exit Sub
	End If
	Dim InsertName,InsertValue
	Do while not Rs.eof
	InsertName=""
	InsertValue=""
	For i = 1 to Rs.Fields.Count-1
		InsertName=InsertName & Rs(i).Name
		InsertValue=InsertValue & "'" & Dvbbs.checkStr(Rs(i)) & "'"
		If i<> Rs.Fields.Count-1 Then 
			InsertName	= InsertName & ","
			InsertValue	= InsertValue & ","
		End If
	Next
	StyleConn.Execute("insert into [Dv_Style] ("&InsertName&") values ("&InsertValue&") ")
	'StyleConn.Execute("Update [Dv_Style] set "&SQLSTR&" where ID="&SkinMdbID)
	Rs.movenext
	loop
	Rs.close
	set Rs=nothing
	Dv_suc(SucMsg&"<li>���ݵ����ɹ���")
End If
End Sub

Sub Load()
%>
<form action="?action=loadthis" method=post>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<tr><th colspan="2">����ģ������</th></tr>
<tr>
<td width="20%" class="forumrow">����ģ�����ݿ�����</td>
<td width="80%" class="forumrow"><input type="text" name="skinmdb" size="30" value="Skins/Dv_skin.mdb"></td>
</tr>
<tr><th colspan="2"><input type="submit" name="submit" value="��һ��"></th></tr>
</table></form>
<%
End Sub

Sub loadskin()
Dim tRs
skid=Dvbbs.checkstr(Request("skid"))
mdbname=Dvbbs.Checkstr(trim(Request.form("skinmdb")))
If skid="" or isnull(skid) or Not Isnumeric(Replace(Replace(skid,",","")," ","")) Then
	Errmsg=ErrMsg + "<BR><li>����δѡȡҪ�����ģ��"
	Exit Sub
End If
If mdbname="" Then
	Errmsg=ErrMsg + "<BR><li>����д����ģ�����ݿ���"
	Exit Sub
End If
SkinConnection(mdbname)
If Request("submit")="ɾ��" Then
	StyleConn.Execute("Delete from Dv_Style where id in ("&skid&")")
	Dv_suc("ɾ���ɹ���")
Else
ChkSkinMDB()
if Errmsg<>"" Then Exit Sub
Dim InsertName,InsertValue
Set TRs=StyleConn.Execute(" select * from Dv_Style where id in ("&skid&")  order by id ")
Do while not TRs.eof
InsertName=""
InsertValue=""
	For i = 1 to TRs.Fields.Count-1
		InsertName=InsertName & TRs(i).Name
		InsertValue=InsertValue & "'" & Dvbbs.checkStr(TRs(i)) & "'"
		If i<> TRs.Fields.Count-1 Then 
			InsertName	= InsertName & ","
			InsertValue	= InsertValue & ","
		End If
	Next
Dvbbs.Execute("insert into [Dv_Style] ("&InsertName&") values ("&InsertValue&") ")
TRs.movenext
loop
TRs.close
set Rs=nothing
set TRs=nothing
Dv_suc("���ݵ���ɹ���")
Dvbbs.DelCahe("Templateslist")
End If

End Sub

'ģ�����
Sub rename()
Dim sRs
skid=Dvbbs.checkstr(Request("skid"))
mdbname=Dvbbs.Checkstr(Trim(Request("mdbname")))
If skid<>"" and IsNumeric(skid) Then skid=Clng(skid) Else skid=1
If Request("act")="loadskin" and mdbname<>"" Then      
	SkinConnection(mdbname)
	set sRs=StyleConn.Execute("select id,StyleName from Dv_Style where id="&skid)
Else
	set sRs=Dvbbs.Execute("select id,StyleName from Dv_Style where id="&skid)
End If
%>
<form action="?action=savenm" method=post >
<table border="0"  cellspacing="1" cellpadding="5" align=center width="95%" class="tableBorder">
<tr><th colspan="2">����ģ������ ID=<%=sRs(0)%></td></tr>
<tr>
	<td width="20%" class="forumrow">ģ��ԭ����</td>
	<td width="80%" class="forumrow"><%=sRs(1)%></td>
</tr>
<tr>
	<td class="forumrow">ģ��������</td>
	<td class="forumrow"><input type="text" name="skinNAME" size="30" value=""></td>
</tr>
<tr><th colspan="2"><input type="submit" name="submit" value="����"></th></tr>
<% If Request("act")="loadskin" Then
%><input TYPE="hidden" NAME="mdbname" VALUE="<% =mdbname %>">
<% End If %>
<input TYPE="hidden" NAME="skid" VALUE="<% =sRs(0) %>">
<input TYPE="hidden" NAME="act" VALUE="<% =Request("act") %>">
</table></form>
<%
sRs.close
set sRs=nothing
End Sub

'ģ���������
Sub savenm()
Dim skinNAME
	skid=Dvbbs.checkstr(Request.Form("skid"))
    mdbname=Dvbbs.Checkstr(trim(Request.Form("mdbname")))
	skinNAME=Dvbbs.Checkstr(trim(Request.Form("skinname")))
If skid="" or Not IsNumeric(skid) Then
	Errmsg=ErrMsg + "<BR><li>��ѡ����ȷ�Ĳ���"
	Exit Sub
End IF
If skinNAME=""  Then
	Errmsg=ErrMsg + "<li>��ģ�����Ʋ���Ϊ�գ�"
	Exit Sub
End IF
If Request("act")="loadskin" and mdbname<>"" Then
	SkinConnection(mdbname)
	StyleConn.Execute("UPDATE Dv_Style set StyleName='"&skinNAME&"'  where id="&skid)
Else
	Dvbbs.Execute("UPDATE Dv_Style set StyleName='"&skinNAME&"'  where id="&skid)
	Dvbbs.DelCahe("Templateslist")
End If

Dv_suc("<li>���ݸ��³ɹ���")
End Sub

Sub ChkSkinMDB()
If IsFoundTable("Dv_Style",1)=False Then
	Errmsg=ErrMsg + "<li>"&mdbname&"���ݿ����Ҳ���ָ�������ݱ����½�������ݱ�"
	Errmsg=ErrMsg + "<li><a href=?action=CreatMdb&mdbname="&mdbname&" >���ھ��½�������ݱ�</a>��"
	Exit Sub
End IF
'�������ֶαȽ�
Dim TempField,TempRs,TempSql,FalseName,LostName
TempField=""
FalseName=""
TempSql="Select top 1 * From [Dv_Style]"
If Request("action")="loadskin" Then
Set TempRs = Dvbbs.Execute(TempSql)
Else
Set TempRs = StyleConn.Execute(TempSql)
End If
	For i= 0 to TempRs.Fields.Count-1
		TempField = TempField & TempRs(i).Name &","
	Next
TempRs.Close
TempField=Lcase(TempField)
If Request("action")="loadskin" Then
Set TempRs = StyleConn.Execute(TempSql)
Else
Set TempRs = Dvbbs.Execute(TempSql)
End If
	For i = 0 to TempRs.Fields.Count-1
		If instr(TempField,Lcase(TempRs(i).Name)) = 0 Then
		FalseName = FalseName & TempRs(i).Name &","
		Else
		TempField = Replace(TempField,Lcase(TempRs(i).Name),"")
		TempField = Replace(TempField,",,",",")
		End If
	Next
TempRs.Close
Set TempRs=Nothing
If Right(FalseName,1)="," Then FalseName=Left(FalseName,Len(FalseName)-1)
If Right(TempField,1)="," Then TempField=Left(TempField,Len(TempField)-1)
If Left(TempField,1)="," Then TempField=Replace(TempField,",","",1,1)
If FalseName<>"" Then
	If Request("action")="loadskin" Then
	Errmsg=ErrMsg + "<li>���ݱ��ж�������ֶΣ� "& FalseName &" ����������ݿ�ṹ����ִ�иղŵĲ�����"
	Else
	Call AddFields(FalseName)
	End If
	'Errmsg=ErrMsg + "<li>���ݱ���ȱ���ֶΣ� "& FalseName &" ����������ݿ�ṹ����ִ�иղŵĲ�����"
End If
If TempField<>"" and Request("action")<>"loadskin" Then
SucMsg=SucMsg+"<li>���ݱ��ж�������ֶΣ� "& TempField &" ������Ե����������ɾ��������ֶΣ�"
SucMsg=SucMsg+"<li><a href=?action=DelFields&fields="&TempField&"&mdbname="&mdbname&"><font color=red>ִ������ɾ�����ֶΣ�</font></a>"
End If
End Sub

Sub DelFields()
Dim Fields,TempFields
Fields=Request.QueryString("fields")
If Request("mdbname")="" Then
	Errmsg=ErrMsg + "<BR><li>��ָ������ģ�����ݿ⡣"
	Exit Sub
Else
	mdbname=Dvbbs.Checkstr(Trim(Request("mdbname")))
End If
If Replace(Fields,",","")="" Then Exit Sub
If not IsObject(StyleConn) Then SkinConnection(mdbname)
TempFields=Split(Fields,",")
For i=0 to Ubound(TempFields)
	IF TempFields(i)<>"" Then
	StyleConn.Execute("alter table [Dv_Style] drop COLUMN "&TempFields(i))
	End If
Next
Dv_suc("<li>"&Fields&"ɾ���ɹ���<li><a href=admin_loadskin.asp>����ģ�������ҳ</a>")
End Sub

Sub AddFields(Fields)
If Replace(Fields,",","")="" Then Exit Sub
Dim TempFields,FieldName,FieldSql,FieldValue
TempFields=Split(Fields,",")
If IsObject(StyleConn) Then
For i=0 to Ubound(TempFields)
	Select case Lcase(TempFields(i))
	Case "stylename"
		FieldValue=TempFields(i)	& "=''"
		FieldSql=TempFields(i)		& " varchar(50) NOT NULL"
	Case "forum_css"
		FieldValue=TempFields(i)	& "='|||@@@|||'"
		FieldSql=TempFields(i)		& " text not Null default '|||@@@|||'"
	Case Else
		FieldValue=TempFields(i)	& "='|||@@@|||@@@|||@@@|||'"
		FieldSql=TempFields(i)		& " text not Null default '|||@@@|||@@@|||@@@|||'"
	End Select
If Request("action")="loadskin" Then
Dvbbs.Execute("alter table [Dv_Style] add "&FieldSql)
Dvbbs.Execute("Update [Dv_Style] Set "&FieldValue)
Else
StyleConn.Execute("alter table [Dv_Style] add "&FieldSql)
StyleConn.Execute("Update [Dv_Style] Set "&FieldValue)
End IF
Next
Else
	Errmsg=ErrMsg + "<li>���ݱ�����δ��������"
End If
End Sub

Sub CreateStyleMdb()
'|||@@@|||					--> Forum_CSS
'|||@@@|||@@@|||@@@|||		--> other
If Request("mdbname")="" Then
	Errmsg=ErrMsg + "<BR><li>��ָ������ģ�����ݿ⡣"
	Exit Sub
Else
	mdbname=Dvbbs.Checkstr(Trim(Request("mdbname")))
End If
Dim CreatStr
CreatStr = "CREATE TABLE Dv_Style (ID int IDENTITY (1, 1) NOT NULL CONSTRAINT PK_Dv_Style PRIMARY KEY,"&_
	"StyleName varchar(50) NOT NULL,"&_
	"Forum_CSS text not Null default '|||@@@|||',"
Set Rs=Dvbbs.Execute("select top 1 * From [Dv_Style] ")
	If Rs.EOF Then
		Errmsg=ErrMsg + "<li>�޷�ȡ��Դģ������"
		Dvbbs_error()
		Exit Sub
	End If
	For i= 3 to Rs.Fields.Count-1
		CreatStr=CreatStr & Rs(i).Name & " text not Null default '|||@@@|||@@@|||@@@|||'"
		If i<> Rs.Fields.Count-1 Then 
			CreatStr=CreatStr & ","
		End If
	Next
CreatStr=CreatStr & ")"
Rs.close:Set Rs=Nothing
SkinConnection(mdbname)
StyleConn.Execute(CreatStr)
Dv_suc("<li>Dv_Style���ݱ�ṹ�����ɹ���<li><a href=admin_loadskin.asp>����ģ�������ҳ</a>")
End Sub

'У���ֶ��Ƿ����
Function IsTruePage(page)
IsTruePage=False
If page<>"" Then 
	page=LCase(Trim(page))
	Dim myRs
	Set MyRs=Dvbbs.Execute("Select top 1 * From [Dv_Style]")
	For i= 2 to MyRs.Fields.Count-1
		If LCase(myRs(i).name)=page Then
			IsTruePage=True
			Exit Function
		End If
	Next
	Set MyRs=Nothing
End If
End Function

'�������ֶαȽ�
Sub ChkFields()
Dim TempField,TempRs,TempSql,FalseName,LostName
TempField=""
TempSql="Select top 1 * From [Dv_Style]"
Set TempRs=StyleConn.Execute(TempSql)
	For i= 0 to TempRs.Fields.Count-1
		TempField = TempField & TempRs(i).Name &","
	Next
TempRs.Close
TempField=Lcase(TempField)
Set TempRs = Dvbbs.Execute(TempSql)
	For i = 0 to TempRs.Fields.Count-1
		If instr(TempField,Lcase(TempRs(i).Name)) = 0 Then
		FalseName = FalseName & TempRs(i).Name &","
		Else
		TempField = Replace(TempField,Lcase(TempRs(i).Name),"")
		TempField = Replace(TempField,",,",",")
		End If
	Next
TempRs.Close
Set TempRs=Nothing
End Sub


'У������Ƿ���ڡ�TableName=������str:0=Ĭ�Ͽ⣬1=����
Function IsFoundTable(TableName,Str)
Dim ChkRs
IsFoundTable=False
If TableName<>"" Then 
TableName=LCase(Trim(TableName))
	If Str=0 Then
	Set ChkRs=Conn.openSchema(20)
	Else
	Set ChkRs=StyleConn.openSchema(20)
	End If
	Do Until ChkRs.EOF
		If ChkRs("TABLE_TYPE")="TABLE" Then
			If Lcase(ChkRs("TABLE_NAME"))=TableName then
				IsFoundTable=True
				Exit Function
			End If
		End If
	ChkRs.movenext
	Loop
	ChkRs.close:Set ChkRs=Nothing
End If
End Function
%>