<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/md5.asp"-->
<!-- #include file="inc/DvADChar.asp" -->
<!-- #include file="inc/myadmin.asp" -->
<script language="JavaScript">
<!--

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall')
       e.checked = form.chkall.checked;
    }
  }
//-->
</script>
<%Head()
	Dim admin_flag
	admin_flag=",16,"
	if not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		dvbbs_error()
	else
		dim body,username2,password2,oldpassword,oldusername,oldadduser,username1
'''''''''''''''
'ȡ���û������Ա������ 2002-12-13
		dim groupsname,titlepic
		set rs=Dvbbs.Execute("select usertitle,grouppic from [dv_UserGroups] where UserGroupID=1 ")
		groupsname=rs(0)
		titlepic=rs(1)
		set rs=nothing

		if request("action")="updat" then
			call update()
			response.write body
		elseif request("action")="del" then
			Call Del()
			response.write body
       	elseif request("action")="pasword" then
			call pasword()
       	elseif request("action")="newpass" then
			call newpass()
			response.write body
		elseif request("action")="add" then
			call addadmin()
		elseif request("action")="edit" then
			call userinfo()
		elseif request("action")="savenew" then
			call savenew()
			response.write body
		else
			call userlist()
		end if
		Footer()
	end if

	sub userlist()
%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
                <tr> 
                  <th height=22 colspan=5>����Ա����(����û������в���)</th>
                </tr>
                <tr align=center> 
                  <td width="30%" height=22 class="forumHeaderBackgroundAlternate"><B>�û���</B></td><td width="25%" class="forumHeaderBackgroundAlternate"><B>�ϴε�¼ʱ��</B></td><td width="15%" class="forumHeaderBackgroundAlternate"><B>�ϴε�½IP</B></td><td width="15%" class="forumHeaderBackgroundAlternate"><B>����</B></td>
                </tr>
<%
	set rs=Dvbbs.Execute("select * from "&admintable&" order by LastLogin desc")
	do while not rs.eof
%>
                <tr> 
                  <td class=forumrow><a href="admin_admin.asp?id=<%=rs("id")%>&action=pasword"><%=rs("username")%></a></td><td class=forumrow><%=rs("LastLogin")%></td><td class=forumrow><%=rs("LastLoginIP")%></td><td class=forumrow><a href="admin_admin.asp?action=del&id=<%=rs("id")%>&name=<%=Rs("adduser")%>" onclick="{if(confirm('ɾ����ù���Ա�����ɽ����̨��\n\nȷ��ɾ����?')){return true;}return false;}">ɾ��</a>&nbsp;&nbsp;<a href="admin_admin.asp?id=<%=rs("id")%>&action=edit">�༭Ȩ��</a></td>
                </tr>
<%
	rs.movenext
	loop
	rs.close
	set rs=nothing
%>
	       </table>
<%
	end sub

Sub Del()
	Dim UserTitle
	Rem ���³�������Ա��ĵȼ����� 2004-4-29 Dvbbs.YangZheng
	Sql = "SELECT Top 1 UserTitle From Dv_UserGroups Where MinArticle > 0 And ParentGID = 4 Order By UserGroupID"
	Set Rs = Dvbbs.Execute(Sql)
	If Rs.Eof And Rs.Bof Then
		UserTitle = "ע���Ա"
	Else
		UserTitle = Rs(0)
	End If
	Dvbbs.Execute("DELETE FROM " & Admintable & " WHERE Id = " & Request("Id"))
	Dvbbs.Execute("UPDATE [Dv_User] SET Usergroupid = 4, UserClass = '" & UserTitle & "' WHERE Username = '" & Replace(Request("name"),"'","") & "'")
	body="<li>����Աɾ���ɹ���"
End Sub

sub pasword()
	set rs=Dvbbs.Execute("select * from "&admintable&" where id="&request("id"))
	oldpassword=rs("password")
	oldadduser=rs("adduser")
  %> 
<form action="?action=newpass" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
               <tr> 
                  <th colspan=2 height=23>����Ա���Ϲ����������޸�
                  </th>
                </tr>
               <tr > 
            <td width="26%" align="right" class=forumrow>��̨��¼���ƣ�</td>
            <td width="74%" class=forumrow>
              <input type=hidden name="oldusername" value="<%=rs("username")%>">
              <input type=text name="username2" value="<%=rs("username")%>">  (����ע������ͬ)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow>��̨��¼���룺</td>
            <td width="74%" class=forumrow>
              <input type="password" name="password2" value="<%=oldpassword%>">  (����ע�����벻ͬ,��Ҫ�޸���ֱ������)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow height=23>ǰ̨�û����ƣ�</td>
            <td width="74%" class=forumrow><%=oldadduser%>
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" class=forumrow> 
              <input type=hidden name="adduser" value="<%=oldadduser%>">
              <input type=hidden name=id value="<%=request("id")%>">
              <input type="submit" name="Submit" value="�� ��">
            </td>
          </tr>
        </table>
        </form>

<%       rs.close
         set rs=nothing
end sub

sub newpass()
	dim passnw,usernw,aduser
	set rs=Dvbbs.Execute("select * from "&admintable&" where id="&request("id"))
	oldpassword=rs("password")
	if request("username2")="" then
		Response.Write "<li>���������Ա���֡�<a href=?>�� <font color=red>����</font> ��</a>"
		exit sub
	else 
		usernw=trim(request("username2"))
	end if
	if request("password2")="" then
		Response.Write "<li>�������������롣<a href=?>�� <font color=red>����</font> ��</a>"
		exit sub
	elseif trim(request("password2"))=oldpassword then
		passnw=request("password2")
	else
		passnw=md5(request("password2"),16)
	end if
	if request("adduser")="" then
		Response.Write"<li>���������Ա���֡�<a href=?>�� <font color=red>����</font> ��</a>"
		exit sub
	else 
		aduser=trim(request("adduser"))
	end if

	set rs=server.createobject("adodb.recordset")
	sql="select * from "&admintable&" where username='"&trim(request("oldusername"))&"'"
	rs.open sql,conn,1,3
	if not rs.eof and not rs.bof then
	rs("username")=usernw
	rs("adduser")=aduser
	rs("password")=passnw
''''''''''''''
'�����û��ĵļ���
        Dvbbs.Execute("update [dv_user] set usergroupid=1,userclass='"&groupsname&"',titlepic='"&titlepic&"' where username='"&trim(request("adduser"))&"'")	'
	body="<li>����Ա���ϸ��³ɹ������ס������Ϣ��<br> ����Ա��"&request("username2")&" <BR> ��   �룺"&request("password2")&" <a href=?>�� <font color=red>����</font> ��</a>"
	rs.update
	End if
	rs.close
	set rs=nothing
end sub


sub addadmin()
%> 
<form action="?action=savenew" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
               <tr> 
                  <th colspan=2 height=23>����Ա��������ӹ���Ա
                  </th>
                </tr>
               <tr > 
            <td width="26%" align="right" class=forumrow>��̨��¼���ƣ�</td>
            <td width="74%" class=forumrow>
              <input type=text name="username2" size=30>  (����ע������ͬ)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow>��̨��¼���룺</td>
            <td width="74%" class=forumrow>
              <input type="password" name="password2" size=33>  (����ע�����벻ͬ)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow height=23>ǰ̨�û����ƣ�</td>
            <td width="74%" class=forumrow><input type=text name="username1" size=30>  (��ѡ����д�������޸�)
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" class=forumrow> 
              <input type="submit" name="Submit" value="�� ��">
            </td>
          </tr>
        </table>
        </form>

<%
end sub

sub savenew()
dim adminuserid
	if request.form("username2")="" then
	body="�������̨��¼�û�����"
	exit sub
	end if
	if request.form("username1")="" then
	body="������ǰ̨��¼�û�����"
	exit sub
	end if
	if request.form("password2")="" then
	body="�������̨��¼���룡"
	exit sub
	end if

	set rs=Dvbbs.Execute("select userid from [dv_user] where username='"&replace(request.form("username1"),"'","")&"'")
	if rs.eof and rs.bof then
	body="��������û�������һ����Ч��ע���û���"
	exit sub
        else
        adminuserid=rs(0)
	end if

	set rs=Dvbbs.Execute("select username from "&admintable&" where username='"&replace(request.form("username2"),"'","")&"'")
	if not (rs.eof and rs.bof) then
	body="��������û����Ѿ��ڹ����û��д��ڣ�"
	exit sub
	end if
	Dvbbs.Execute("update [dv_user] set usergroupid=1 , userclass='"&groupsname&"',titlepic='"&titlepic&"' where userid="&adminuserid&" ")
	Dvbbs.Execute("insert into "&Admintable&" (username,[password],adduser) values ('"&replace(request.form("username2"),"'","")&"','"&md5(replace(request.form("password2"),"'",""),16)&"','"&replace(request.form("username1"),"'","")&"')")
	body="�û�ID:"&adminuserid&" ��ӳɹ������ס�¹���Ա��̨��¼��Ϣ�������޸��뷵�ع���Ա����"
end sub

sub userinfo()
dim menu(8,10),trs,k
menu(0,0)="�������"
menu(0,1)="<a href=admin_setting.asp target=main>��������</a>@@1"
menu(0,2)="<a href=admin_ads.asp target=main>������</a>@@2"
menu(0,3)="<a href=admin_log.asp target=main>��̳��־</a>@@3"
menu(0,4)="<a href=admin_help.asp target=main>��������</a>@@4"
menu(0,5)="<a href=admin_wealth.asp target=main>��������</a>@@5"
menu(0,6)="<a href=admin_message.asp target=main>���Ź���</a>@@6"
menu(0,7)="<a href=announcements.asp?boardid=0&action=AddAnn target=_blank>�������</a>@@7"
menu(0,8)="<a href=admin_menpai.asp target=main>���ɹ���</a>@@8"

menu(1,0)="��̳����"
menu(1,1)="<a href=admin_board.asp?action=add target=main>����(����)���</a> | <a href=admin_board.asp target=main>����</a>@@9"
menu(1,2)="<a href=admin_board.asp?action=permission target=main>�ְ����û�Ȩ������</a>@@10"
menu(1,3)="<a href=admin_boardunite.asp target=main>�ϲ���������</a>@@11"
menu(1,4)="<a href=admin_update.asp target=main>�ؼ���̳���ݺ��޸�</a>@@12"
menu(1,5)="<a href=admin_link.asp?action=add target=main>������̳���</a> | <a href=admin_link.asp target=main>����</a>@@13"

menu(2,0)="�û�����"
menu(2,1)="<a href=admin_user.asp target=main>�û�����(Ȩ��)����</a>@@14"
menu(2,2)="<a href=admin_group.asp?action=addgroup target=main>�û������</a> | <a href=admin_group.asp target=main>����</a>@@15"
menu(2,3)="<a href=admin_admin.asp?action=add target=main>����Ա���</a> | <a href=admin_admin.asp target=main>����</a>@@16"
menu(2,4)="<a href=admin_grade.asp?action=add target=main>�û��ȼ����</a> | <a href=admin_grade.asp target=main>����</a>@@17"
menu(2,5)="<a href=admin_update.asp?action=updateuser target=main>�ؼ��û���������</a>@@19"

menu(3,0)="�������"
menu(3,1)="<a href=admin_template.asp target=main>������ģ���ܹ���</a>@@20"
menu(3,2)="<a href=admin_loadskin.asp target=main>ģ�嵼��</a> | <a href=admin_loadskin.asp?action=load target=main>����</a>@@21"

menu(4,0)="��̳���ӹ���"
menu(4,1)="<a href=admin_alldel.asp target=main>����ɾ��</a> | <a href=admin_alldel.asp?action=moveinfo target=main>�����ƶ�</a>@@22"
menu(4,2)="<a href=recycle.asp target=_blank>����վ����</a>@@23"
menu(4,3)="<a href=admin_postdata.asp?action=Nowused target=main>��ǰ�������ݱ����</a>@@24"
menu(4,4)="<a href=admin_postdata.asp target=main>���ݱ������ת��</a>@@25"

menu(5,0)="�滻/���ƴ���"
menu(5,1)="<a href=admin_badword.asp?reaction=badword target=main>�໰��������</a>@@26"
menu(5,2)="<a href=admin_badword.asp?reaction=splitreg target=main>ע������ַ�</a>@@27"
menu(5,3)="<a href=admin_lockip.asp?action=add target=main>IP�����޶����</a> | <a href=admin_lockip.asp target=main>����</a>@@28"
menu(5,4)="<a href=admin_address.asp?action=add target=main>��̳IP�����</a> | <a href=admin_address.asp target=main>����</a>@@29"

menu(6,0)="���ݴ���(Access)"
menu(6,1)="<a href=admin_data.asp?action=CompressData target=main>ѹ�����ݿ�</a>@@30"
menu(6,2)="<a href=admin_data.asp?action=BackupData target=main>�������ݿ�</a>@@31"
menu(6,3)="<a href=admin_data.asp?action=RestoreData target=main>�ָ����ݿ�</a>@@32"
menu(6,4)="<a href=admin_data.asp?action=SpaceSize target=main>ϵͳ�ռ�ռ��</a>@@33"

menu(7,0)="�ļ�����"
menu(7,1)="<a href=admin_upUserface.asp target=main>�ϴ�ͷ�����</a>@@34"
menu(7,2)="<a href=admin_uploadlist.asp target=main>�ϴ��ļ�����</a>@@35"

menu(8,0)="�˵�����"
menu(8,1)="<a href=admin_plus.asp target=main>��̳�˵�����</a>@@36"

dim j,tmpmenu,menuname,menurl
set rs=Dvbbs.Execute("select * from "&admintable&" where id="&request("id"))
%>
<form action="admin_admin.asp?action=updat" method=post name=adminflag>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr> 
<th height=25><b>����ԱȨ�޹���</b>(��ѡ����Ӧ��Ȩ�޷��������Ա <%=rs("username")%>)
</th>
</tr>
<tr> 
<td height=25 class="forumHeaderBackgroundAlternate"><b>>>ȫ��Ȩ��</b></td></tr>
<tr><td class=forumrow>
<%for i=0 to ubound(menu,1)%>
<b><%=menu(i,0)%></b><br>
<%
on error resume next
for j=1 to ubound(menu,2)
if isempty(menu(i,j)) then exit for
tmpmenu=split(menu(i,j),"@@")
menuname=tmpmenu(0)
menurl=tmpmenu(1)
%>
<input type="checkbox" name="flag" <% if instr(","&session("flag")&",",",16,")=0 then response.write "disabled=true" %> value="<%=menurl%>" <% if instr(","&rs("flag")&",",","&menurl&",")>0 then response.write "checked" %>><%=menurl%>.<%=menuname%>&nbsp;&nbsp;
<%next%><br><br>
<%next%>
<input type=hidden name=id value="<%=request("id")%>">
<input type="submit" name="Submit" value="����">������<input name=chkall type=checkbox value=on onclick=CheckAll(this.form)>ѡ������Ȩ��
</td>
</tr>
</table>
</form>
<%
rs.close
set rs=nothing
end sub

sub update()
' 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35
'Response.Write request("flag")
'response.end
set rs=server.createobject("adodb.recordset")
sql="select * from "&admintable&" where id="&request("id")
rs.open sql,conn,1,3
if not rs.eof and not rs.bof then
rs("flag")=replace(request("flag")," ","")
body="<li>����Ա���³ɹ������ס������Ϣ��"
rs.update
if rs("adduser")=Dvbbs.membername then session("flag")=replace(request("flag")," ","")
end if
rs.close
set rs=nothing
end sub

%>