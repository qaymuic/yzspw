<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
Head()
dim action
dim admin_flag
action=trim(request("action"))

dim dbpath,bkfolder,bkdbname,fso,fso1
Dim uploadpath
If Dvbbs.Forum_Setting(76)="0" Or  Dvbbs.Forum_Setting(76)="" Then Dvbbs.Forum_Setting(76)="UploadFile/"
uploadpath=Dvbbs.Forum_Setting(76)
select case action
case "CompressData"		'ѹ������
	admin_flag=",30,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		dim tmprs
		dim allarticle
		dim Maxid
		dim topic,username,dateandtime,body
		
		call CompressData()

	end if	

case "BackupData"		'��������
	admin_flag=",31,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		if request("act")="Backup" then
		call updata()
		else
		
		call BackupData()
		end if
	end if

case "RestoreData"		'�ָ�����
	admin_flag=",32,"
	dim backpath
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		if request("act")="Restore" then
			Dbpath=request.form("Dbpath")
			backpath=request.form("backpath")
			if dbpath="" then
			response.write "��������Ҫ�ָ��ɵ����ݿ�ȫ��"	
			else
			Dbpath=server.mappath(Dbpath)
			end if
			backpath=server.mappath(backpath)
		
			Set Fso=server.createobject("scripting.filesystemobject")
			if fso.fileexists(dbpath) then  					
			fso.copyfile Dbpath,Backpath
			response.write "�ɹ��ָ����ݣ�"
			else
			response.write "����Ŀ¼�²������ı����ļ���"	
			end if
		else
		
		call RestoreData()
		end if
	end if

 case "SpaceSize"		'ϵͳ�ռ�ռ��
	admin_flag=",33,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li> ��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		call SpaceSize()
	end if

case else
		Errmsg=ErrMsg + "<BR><li>ѡȡ��Ӧ�Ĳ�����"
		dvbbs_error()

end select

Footer()
response.write"</body></html>"


'====================ϵͳ�ռ�ռ��=======================
sub SpaceSize()
On error resume next
%>
		<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">			<tr>
  					<th height=25>
  					&nbsp;&nbsp;ϵͳ�ռ�ռ�����
  					</th>
  				</tr> 	
 				<tr>
 					<td class="forumrow"> 			
 			<blockquote>
 			<br> 			
 			��������ռ�ÿռ䣺&nbsp;<img src="skins/default/bar/bar1.gif" width=<%=drawbar("data")%> height=10>&nbsp;<%showSpaceinfo("data")%><br><br>
 			��������ռ�ÿռ䣺&nbsp;<img src="skins/default/bar/bar1.gif" width=<%=drawbar("databackup")%> height=10>&nbsp;<%showSpaceinfo("databackup")%><br><br>
 			�����ļ�ռ�ÿռ䣺&nbsp;<img src="skins/default/bar/bar1.gif" width=<%=drawspecialbar%> height=10>&nbsp;<%showSpecialSpaceinfo("Program")%><br><br>
 			ģ��ͼƬռ�ÿռ䣺&nbsp;<img src="skins/default/bar/bar1.gif" width=<%=drawbar("skins")%> height=10>&nbsp;<%showSpaceinfo("skins")%><br><br>
 			ϵͳͼƬռ�ÿռ䣺&nbsp;<img src="skins/default/bar/bar1.gif" width=<%=drawbar("images")%> height=10>&nbsp;<%showSpaceinfo("images")%><br><br>
 			�ϴ�ͷ��ռ�ÿռ䣺&nbsp;<img src="skins/default/bar/bar1.gif" width=<%=drawbar("uploadFace")%> height=10>&nbsp;<%showSpaceinfo("uploadFace")%><br><br>
 			<%
 			Dim tmpstr
 			tmpstr=uploadpath
 			%>
 			�ϴ�ͼƬռ�ÿռ䣺&nbsp;<img src="skins/default/bar/bar1.gif" width=<%=drawbar(tmpstr)%> height=10>&nbsp;<%showSpaceinfo(uploadpath)%><br><br>	
 			ϵͳռ�ÿռ��ܼƣ�<br><img src="skins/default/bar/bar1.gif" width=400 height=10> <%showspecialspaceinfo("All")%>
 			</blockquote> 	
 					</td>
 				</tr>
 			</table>
<%
end sub

sub SQLUserReadme()
%>
		<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">			<tr>
  					<th height=25>
  					&nbsp;&nbsp;SQL���ݿ����ݴ���˵��
  					</th>
  				</tr> 	
 				<tr>
 					<td class="forumrow"> 			
 			<blockquote>
<B>һ���������ݿ�</B>
<BR><BR>
1����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server<BR>
2��SQL Server��-->˫������ķ�����-->˫�������ݿ�Ŀ¼<BR>
3��ѡ��������ݿ����ƣ�����̳���ݿ�Forum��-->Ȼ�������˵��еĹ���-->ѡ�񱸷����ݿ�<BR>
4������ѡ��ѡ����ȫ���ݣ�Ŀ���еı��ݵ����ԭ����·����������ѡ�����Ƶ�ɾ����Ȼ�����ӣ����ԭ��û��·����������ֱ��ѡ����ӣ�����ָ��·�����ļ�����ָ�����ȷ�����ر��ݴ��ڣ����ŵ�ȷ�����б���
<BR><BR>
<B>������ԭ���ݿ�</B><BR><BR>
1����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server<BR>
2��SQL Server��-->˫������ķ�����-->��ͼ�������½����ݿ�ͼ�꣬�½����ݿ����������ȡ<BR>
3������½��õ����ݿ����ƣ�����̳���ݿ�Forum��-->Ȼ�������˵��еĹ���-->ѡ��ָ����ݿ�<BR>
4���ڵ������Ĵ����еĻ�ԭѡ����ѡ����豸-->��ѡ���豸-->�����-->Ȼ��ѡ����ı����ļ���-->��Ӻ��ȷ�����أ���ʱ���豸��Ӧ�ó������ղ�ѡ������ݿⱸ���ļ��������ݺ�Ĭ��Ϊ1���������ͬһ���ļ�������α��ݣ����Ե�����ݺ��ԱߵĲ鿴���ݣ��ڸ�ѡ����ѡ�����µ�һ�α��ݺ��ȷ����-->Ȼ�����Ϸ������Աߵ�ѡ�ť<BR>
5���ڳ��ֵĴ�����ѡ�����������ݿ���ǿ�ƻ�ԭ���Լ��ڻָ����״̬��ѡ��ʹ���ݿ���Լ������е��޷���ԭ����������־��ѡ��ڴ��ڵ��м䲿λ�Ľ����ݿ��ļ���ԭΪ����Ҫ������SQL�İ�װ�������ã�Ҳ����ָ���Լ���Ŀ¼�����߼��ļ�������Ҫ�Ķ������������ļ���Ҫ���������ָ��Ļ���������Ķ���������SQL���ݿ�װ��D:\Program Files\Microsoft SQL Server\MSSQL\Data����ô�Ͱ������ָ�������Ŀ¼������ظĶ��Ķ������������ļ�����øĳ�����ǰ�����ݿ�������ԭ����bbs_data.mdf�����ڵ����ݿ���forum���͸ĳ�forum_data.mdf������־�������ļ���Ҫ���������ķ�ʽ����صĸĶ�����־���ļ�����*_log.ldf��β�ģ�������Ļָ�Ŀ¼�������������ã�ǰ���Ǹ�Ŀ¼������ڣ���������ָ��d:\sqldata\bbs_data.mdf����d:\sqldata\bbs_log.ldf��������ָ�������<BR>
6���޸���ɺ󣬵�������ȷ�����лָ�����ʱ�����һ������������ʾ�ָ��Ľ��ȣ��ָ���ɺ�ϵͳ���Զ���ʾ�ɹ������м���ʾ�������¼����صĴ������ݲ�ѯ�ʶ�SQL�����Ƚ���Ϥ����Ա��һ��Ĵ����޷���Ŀ¼��������ļ����ظ������ļ���������߿ռ䲻���������ݿ�����ʹ���еĴ������ݿ�����ʹ�õĴ��������Գ��Թر����й���SQL����Ȼ�����´򿪽��лָ��������������ʾ����ʹ�õĴ�����Խ�SQL����ֹͣȻ�����𿴿����������������Ĵ���һ�㶼�ܰ��մ�����������Ӧ�Ķ��󼴿ɻָ�<BR><BR>

<B>�����������ݿ�</B><BR><BR>
һ������£�SQL���ݿ�����������ܴܺ�̶��ϼ�С���ݿ��С������Ҫ������������־��С��Ӧ�����ڽ��д˲����������ݿ���־����<BR>
1���������ݿ�ģʽΪ��ģʽ����SQL��ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����-->˫�������ݿ�Ŀ¼-->ѡ��������ݿ����ƣ�����̳���ݿ�Forum��-->Ȼ�����Ҽ�ѡ������-->ѡ��ѡ��-->�ڹ��ϻ�ԭ��ģʽ��ѡ�񡰼򵥡���Ȼ��ȷ������<BR>
2���ڵ�ǰ���ݿ��ϵ��Ҽ��������������е��������ݿ⣬һ�������Ĭ�����ò��õ�����ֱ�ӵ�ȷ��<BR>
3��<font color=blue>�������ݿ���ɺ󣬽��齫�������ݿ�������������Ϊ��׼ģʽ����������ͬ��һ�㣬��Ϊ��־��һЩ�쳣����������ǻָ����ݿ����Ҫ����</font>
<BR><BR>

<B>�ġ��趨ÿ���Զ��������ݿ�</B><BR><BR>
<font color=red>ǿ�ҽ������������û����д˲�����</font><BR>
1������ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����<BR>
2��Ȼ�������˵��еĹ���-->ѡ�����ݿ�ά���ƻ���<BR>
3����һ��ѡ��Ҫ�����Զ����ݵ�����-->��һ�����������Ż���Ϣ������һ�㲻����ѡ��-->��һ��������������ԣ�Ҳһ�㲻ѡ��<BR>
4����һ��ָ�����ݿ�ά���ƻ���Ĭ�ϵ���1�ܱ���һ�Σ��������ѡ��ÿ�챸�ݺ��ȷ��<BR>
5����һ��ָ�����ݵĴ���Ŀ¼��ѡ��ָ��Ŀ¼������������D���½�һ��Ŀ¼�磺d:\databak��Ȼ��������ѡ��ʹ�ô�Ŀ¼������������ݿ�Ƚ϶����ѡ��Ϊÿ�����ݿ⽨����Ŀ¼��Ȼ��ѡ��ɾ�����ڶ�����ǰ�ı��ݣ�һ���趨4��7�죬�⿴���ľ��屸��Ҫ�󣬱����ļ���չ��һ�㶼��bak����Ĭ�ϵ�<BR>
6����һ��ָ��������־���ݼƻ�����������Ҫ��ѡ��-->��һ��Ҫ���ɵı���һ�㲻��ѡ��-->��һ��ά���ƻ���ʷ��¼�������Ĭ�ϵ�ѡ��-->��һ�����<BR>
7����ɺ�ϵͳ�ܿ��ܻ���ʾSql Server Agent����δ�������ȵ�ȷ����ɼƻ��趨��Ȼ���ҵ��������ұ�״̬���е�SQL��ɫͼ�꣬˫���㿪���ڷ�����ѡ��Sql Server Agent��Ȼ�������м�ͷ��ѡ���·��ĵ�����OSʱ�Զ���������<BR>
8�����ʱ�����ݿ�ƻ��Ѿ��ɹ��������ˣ�������������������ý����Զ�����
<BR><BR>
�޸ļƻ���<BR>
1������ҵ���������ڿ���̨��Ŀ¼�����ε㿪Microsoft SQL Server-->SQL Server��-->˫������ķ�����-->����-->���ݿ�ά���ƻ�-->�򿪺�ɿ������趨�ļƻ������Խ����޸Ļ���ɾ������
<BR><BR>
<B>�塢���ݵ�ת�ƣ��½����ݿ��ת�Ʒ�������</B><BR><BR>
һ������£����ʹ�ñ��ݺͻ�ԭ����������ת�����ݣ�����������£������õ��뵼���ķ�ʽ����ת�ƣ�������ܵľ��ǵ��뵼����ʽ�����뵼����ʽת������һ�����þ��ǿ������������ݿ���Ч�������������С�����������ݿ�Ĵ�С��������Ĭ��Ϊ����SQL�Ĳ�����һ�����˽⣬��������еĲ��ֲ�������⣬������ѯ���������Ա���߲�ѯ��������<BR>
1����ԭ���ݿ�����б��洢���̵�����һ��SQL�ļ���������ʱ��ע����ѡ����ѡ���д�����ű��ͱ�д�����������Ĭ��ֵ�ͼ��Լ���ű�ѡ��<BR>
2���½����ݿ⣬���½����ݿ�ִ�е�һ������������SQL�ļ�<BR>
3����SQL�ĵ��뵼����ʽ���������ݿ⵼��ԭ���ݿ��е����б�����<BR>
 			</blockquote> 	
 					</td>
 				</tr>
 			</table>
<%
end sub

'====================�ָ����ݿ�=========================
sub RestoreData()
If IsSqlDataBase = 1 Then
	SQLUserReadme()
	Exit Sub
End If
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder"		<tr>
	<th height=25 >
   					&nbsp;&nbsp;<B>�ָ���̳����</B>( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )
  					</th>
  				</tr>
				<form method="post" action="ADMIN_data.asp?action=RestoreData&act=Restore">
  				
  				<tr>
  					<td height=100 class="forumrow">
  						&nbsp;&nbsp;�������ݿ�·��(���)��<input type=text size=30 name=DBpath value="DataBackup\dvbbs7_Backup.MDB">&nbsp;&nbsp;<BR>
  						&nbsp;&nbsp;Ŀ�����ݿ�·��(���)��<input type=text size=30 name=backpath value="<%=db%>"><BR>&nbsp;&nbsp;��д����ǰʹ�õ����ݿ�·�����粻�븲�ǵ�ǰ�ļ���������������ע��·���Ƿ���ȷ����Ȼ���޸�conn.asp�ļ������Ŀ���ļ����͵�ǰʹ�����ݿ���һ�µĻ��������޸�conn.asp�ļ�<BR>
						&nbsp;&nbsp;<input type=submit value="�ָ�����"> <br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;��������д����������ݿ�·��ȫ�����������Ĭ�ϱ������ݿ��ļ�ΪDataBackup\dvbbs_Backup.MDB���밴�����ı����ļ������޸ġ�<br>
  						&nbsp;&nbsp;������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
  						&nbsp;&nbsp;ע�⣺����·��������������ռ��Ŀ¼�����·��</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end sub

'====================�������ݿ�=========================
sub BackupData()
If IsSqlDataBase = 1 Then
	SQLUserReadme()
	Exit Sub
End If
%>
	<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  				<tr>
  					<th height=25 >
  					&nbsp;&nbsp;<B>������̳����</B>( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )
  					</th>
  				</tr>
  				<form method="post" action="ADMIN_data.asp?action=BackupData&act=Backup">
  				<tr>
  					<td height=100 class="forumrow">
  						&nbsp;&nbsp;
						��ǰ���ݿ�·��(���·��)��<input type=text size=15 name=DBpath value="<%=db%>"><BR>&nbsp;&nbsp;
						�������ݿ�Ŀ¼(���·��)��<input type=text size=15 name=bkfolder value=Databackup>&nbsp;��Ŀ¼�����ڣ������Զ�����<BR>&nbsp;&nbsp;
						�������ݿ�����(��д����)��<input type=text size=15 name=bkDBname value=dvbbs7.MDB>&nbsp;�籸��Ŀ¼�и��ļ��������ǣ���û�У����Զ�����<BR>
						&nbsp;&nbsp;<input type=submit value="ȷ��"><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;��������д����������ݿ�·��ȫ�����������Ĭ�����ݿ��ļ�ΪData\dvbbs7.MDB��<B>��һ��������Ĭ�����������������ݿ�</B><br>
  						&nbsp;&nbsp;������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
  						&nbsp;&nbsp;ע�⣺����·��������������ռ��Ŀ¼�����·��				</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end sub

sub updata()
		Dbpath=request.form("Dbpath")
		Dbpath=server.mappath(Dbpath)
		bkfolder=request.form("bkfolder")
		bkdbname=request.form("bkdbname")
		Set Fso=server.createobject("scripting.filesystemobject")
		if fso.fileexists(dbpath) then
			If CheckDir(bkfolder) = True Then
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			else
			MakeNewsDir bkfolder
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			end if
			response.write "�������ݿ�ɹ��������ݵ����ݿ�·��Ϊ" &bkfolder& "\"& bkdbname
		Else
			response.write "�Ҳ���������Ҫ���ݵ��ļ���"
		End if
end sub
'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '����
       CheckDir = True
    Else
       '������
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------����ָ����������Ŀ¼-----------------------
Function MakeNewsDir(foldername)
	dim f
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function


'====================ѹ�����ݿ� =========================
sub CompressData()

If IsSqlDataBase = 1 Then
	SQLUserReadme()
	Exit Sub
End If
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<form action="Admin_data.asp?action=CompressData" method="post">
<tr>
<td class="forumrow" height=25><b>ע�⣺</b><br>�������ݿ��������·��,�����������ݿ����ƣ�����ʹ�������ݿⲻ��ѹ������ѡ�񱸷����ݿ����ѹ�������� </td>
</tr>
<tr>
<td class="forumrow">ѹ�����ݿ⣺<input type="text" name="dbpath" value=Data\dvbbs7.MDB>&nbsp;
<input type="submit" value="��ʼѹ��"></td>
</tr>
<tr>
<td class="forumrow"><input type="checkbox" name="boolIs97" value="True">���ʹ�� Access 97 ���ݿ���ѡ��
(Ĭ��Ϊ Access 2000 ���ݿ�)<br><br></td>
</tr>
<form>
</table>
<%
dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If

end sub

'=====================ѹ������=========================
Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath,JET_3X
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dbPath) Then
	fso.CopyFile dbpath,strDBPath & "temp.mdb"
	Set Engine = CreateObject("JRO.JetEngine")

	If boolIs97 = "True" Then
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb", _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp1.mdb"
	End If

fso.CopyFile strDBPath & "temp1.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
fso.DeleteFile(strDBPath & "temp1.mdb")
Set fso = nothing
Set Engine = nothing

	CompactDB = "������ݿ�, " & dbpath & ", �Ѿ�ѹ���ɹ�!" & vbCrLf

Else
	CompactDB = "���ݿ����ƻ�·������ȷ. ������!" & vbCrLf
End If

End Function


'=====================ϵͳ�ռ����=========================
	Sub ShowSpaceInfo(drvpath)
 		dim fso,d,size,showsize
 		set fso=server.createobject("scripting.filesystemobject") 		
 		drvpath=server.mappath(drvpath) 		 		
 		set d=fso.getfolder(drvpath) 		
 		size=d.size
 		showsize=size & "&nbsp;Byte" 
 		if size>1024 then
 		   size=(Size/1024)
 		   showsize=size & "&nbsp;KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;MB"		
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;GB"	   
 		end if   
 		response.write "<font face=verdana>" & showsize & "</font>"
 	End Sub	
 	
 	Sub Showspecialspaceinfo(method)
 		dim fso,d,fc,f1,size,showsize,drvpath 		
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpath=server.mappath("pic")
 		drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
 		set d=fso.getfolder(drvpath) 		
 		
 		if method="All" then 		
 			size=d.size
 		elseif method="Program" then
 			set fc=d.Files
 			for each f1 in fc
 				size=size+f1.size
 			next	
 		end if	
 		
 		showsize=size & "&nbsp;Byte" 
 		if size>1024 then
 		   size=(Size/1024)
 		   showsize=size & "&nbsp;KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;MB"		
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;GB"	   
 		end if   
 		response.write "<font face=verdana>" & showsize & "</font>"
 	end sub 	 	 	
 	
 	Function Drawbar(drvpath)
 		dim fso,drvpathroot,d,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath("pic")
 		drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		
 		drvpath=server.mappath(drvpath) 		
 		set d=fso.getfolder(drvpath)
 		size=d.size
 		
 		barsize=cint((size/totalsize)*400)
 		Drawbar=barsize
 	End Function 	
 	
 	Function Drawspecialbar()
 		dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath("pic")
 		drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		
 		set fc=d.files
 		for each f1 in fc
 			size=size+f1.size
 		next	
 		
 		barsize=cint((size/totalsize)*400)
 		Drawspecialbar=barsize
 	End Function 	
%>