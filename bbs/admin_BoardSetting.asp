<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%
	Head()
	Dim Board_Setting
	if not Dvbbs.master or instr(","&session("flag")&",",",9,")=0 then
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		if request("action")="save" then
		call saveconst()
		else
		call consted()
		end if
		Footer()
	end if

sub consted()
if not isnumeric(request("editid")) then
	Errmsg=ErrMsg + "<BR><li>����İ�����Ϣ"
	dvbbs_error()
	exit sub
end if
set rs=Dvbbs.Execute("select * from dv_board where boardid="&request("editid"))
Board_Setting=split(rs("board_setting"),",")
%>
<table width="95%" cellspacing="1" cellpadding="1"  align=center class="tableBorder">
<tr><th height="25" colspan="5" align=left>��̳�߼����� �� <%=rs("boardtype")%></th></tr>
<tr> 
<td width="100%" class=Forumrow colspan=5 height=25>
˵����<BR>
1������ϸ��������ĸ߼�ѡ�Flash��ǩ����򿪣��԰�ȫ��һ��Ӱ�죬��������ľ���������ǡ�<BR>
2�������Խ��߼����õ�ĳ�����ã�ѡ����������ұߵĸ�ѡ�򣩱��浽���а��桢��ͬ���������а��棨���������ࣩ����ͬ���������а��棨�������ࣩ��ͬ����ͬ������棬�������������ز�����<BR>
3��<font color=red>ע�⣬ѡ���������°������⽫��ʹ����ͬ����</font>��
</td>
</tr>
<form method="POST" action="admin_boardsetting.asp?action=save">
<input type=hidden value="<%=request("editid")%>" name="editid">
<tr> 
<td width="100%" class=ForumrowHighlight colspan=5 height=25>
<font color=blue>����Ŀ��</font>��<input type=radio name="savetype" value=0 checked>�ð���&nbsp;<input type=radio name="savetype" value=1>���а���&nbsp;<input type=radio name="savetype" value=2>��ͬ���������а��棨���������ࣩ&nbsp;<input type=radio name="savetype" value=3>��ͬ���������а��棨�������ࣩ&nbsp;<input type=radio name="savetype" value=4>ͬ����ͬ�������
</td>
</tr>
<tr> 
<td width="100%" class=Forumrow colspan=5 height=25>
<font color=blue>
����ָ�ķ����ָһ�����࣬�����Ǹð�����ϼ�����</font>��������Ŀǰ���õ���һ���弶���棬ѡ������ͬ���������а��涼���£���ô���ｫ���°����÷����һ�����������������ļ����а��棬��������ĸ��·�Χ̫�󣬿���ѡ�����ͬ����ͬ������档
</td>
</tr>
<tr><th height="25" colspan="5" align=left> &nbsp;�������õ���</th></tr>
<tr> 
<td width="100%" class=Forumrow colspan=5 height=25>
[<a href="#setting1">��������</a>]
[<a href="#setting2">����Ȩ��</a>]
[<a href="#setting3">ǰ̨����Ȩ��</a>]
[<a href="#setting4">�������</a>]
[<a href="#setting5">�����б���ʾ</a>]
[<a href="#setting6">����������ʾ</a>]
[<a href="#setting7">������������</a>]
[<a href="#setting8">��̳ר������</a>]
[<a href="#setting9">��̳������������</a>]
</td>
</tr>

<tr><th height="25" colspan="5" id=tabletitlelink align=left> &nbsp;<a name="setting1">��������</a>[<a href="#top">����</a>]</th></tr>
<tr> 
<td width="50%" colspan=2 class=Forumrow>
<U>�ⲿ����</U><BR>��д�����ݺ�����̳�б����˰��潫�Զ��л�������ַ<BR>����дURL����·��</td>
<td colspan=2 class=Forumrow>
<input type=text name="Board_Setting(50)" value="<%=Board_Setting(50)%>" size=50>
</td>
<input type="hidden" id="b0" value="<b>�ⲿ����</b><br><li>��д�����ݺ�����̳�б����˰��潫�Զ��л�������ַ<br><li>����дURL����·��">
<td class=Forumrow><a href=# onclick="helpscript(b0);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td width="50%" colspan=2 class=ForumrowHighlight>
<U>����̳LOGO</U><BR>��дͼƬ����Ի����·��������д��ǰ����LOGOΪ��̳������LOGO</td>
<td colspan=2 class=ForumrowHighlight>
<input type=text name="Board_Setting(51)" value="<%=Board_Setting(51)%>" size=50>
</td>
<input type="hidden" id="ba1" value="<b>����̳LOGO</b><br><li>��дͼƬ����Ի����·��������д��ǰ����LOGOΪ��̳������LOGO">
<td class=ForumrowHighlight><a href=# onclick="helpscript(ba1);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�Ƿ���ð����̳��ƶ�</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(40)" value=0 <%if Board_Setting(40)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(40)" value=1 <%if Board_Setting(40)="1" then%>checked<%end if%>>����&nbsp;
</td>
<input type="hidden" id="b6" value="<b>�Ƿ���ð����̳��ƶ�</b><br><li>������ø��ƶȣ����ϼ���̳�����ɹ����¼���̳�����Ϣ">
<td class=Forumrow><a href=# onclick="helpscript(b6);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>��̳�б���ʾ������̳���</U><BR></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(39)" value=0 <%if Board_Setting(39)="0" then%>checked<%end if%>>�б�&nbsp;
<input type=radio name="Board_Setting(39)" value=1 <%if Board_Setting(39)="1" then%>checked<%end if%>>���&nbsp;
</td>
<input type="hidden" id="b7" value="<b>��̳�б���ʾ������̳���</b><br><li>������̳��������̳��ʱ����Ч">
<td class=ForumrowHighlight><a href=# onclick="helpscript(b7);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>��̳�б�����һ�а�����</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(41)" value="<%=Board_Setting(41)%>"> ��
</td>
<input type="hidden" id="b8" value="<b>��̳�б�����һ�а�����</b><br><li>����̳�б�����������̳���Ϊ��࣬��ѡ����Ч����ѡ��Ϊ���ü����̳�б���һ�����а�����">
<td class=Forumrow><a href=# onclick="helpscript(b8);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�Ƿ񹫿���̳�¼��еĲ�����</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(36)" value=0 <%if Board_Setting(36)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(36)" value=1 <%if Board_Setting(36)="1" then%>checked<%end if%>>��&nbsp;
</td>
<input type="hidden" id="b12" value="<b>�Ƿ񹫿���̳�¼��еĲ�����</b><br><li>��̳�ж����ӵ�ɾ�����̶������þ����Ȳ�������Ҫ��¼�����ߺͲ������ݵģ�����ԱĬ�Ͽɿ�����Щ�������ݣ�һ���û�������˴�ѡ����ǽ��ܿ���������">
<td class=Forumrow><a href=# onclick="helpscript(b12);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr><th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting2">����Ȩ�����</a>[<a href="#top">����</a>]</th></tr>
<tr> 
<td width="50%" colspan=2 class=Forumrow>
<U>����̳��Ϊ������̳��������</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(43)" value=0 <%if Board_Setting(43)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(43)" value=1 <%if Board_Setting(43)="1" then%>checked<%end if%>>��&nbsp;
</td>
<input type="hidden" id="b1" value="<b>����̳��Ϊ������̳��������</b><br><li>����Ѿ���������ʾ����������ת�Ƶ������̳<br><li>ѡ���˸�������л�Ա�������ڱ��淢��/�����Ȳ���">
<td class=Forumrow><a href=# onclick="helpscript(b1);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�Ƿ�������̳</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(0)" value=0 <%if Board_Setting(0)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(0)" value=1 <%If Board_Setting(0)="1" then%>checked<%end if%>>��&nbsp;
</td>
<input type="hidden" id="b2" value="<b>�Ƿ�������̳</b><br><li>������ֻ̳�й���Ա�͸ð�������ɽ�">
<td class=ForumrowHighlight><a href=# onclick="helpscript(b2);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�Ƿ�������̳</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(1)" value=0 <%If Board_Setting(1)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(1)" value=1 <%if Board_Setting(1)="1" then%>checked<%end if%>>��&nbsp;
</td>
<input type="hidden" id="b3" value="<b>�Ƿ�������̳</b><br><li>������ֻ̳�й���Ա�͸ð�������ɼ��ͽ���<br><li>����û������̳Ȩ�޹�����û�Ȩ�޹������������û��ɼ��ͽ���<br><li>�����ƶ�һ����̳����Ч">
<td class=Forumrow><a href=# onclick="helpscript(b3);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�Ƿ���֤��̳</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(2)" value=0 <%if Board_Setting(2)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(2)" value=1 <%if Board_Setting(2)="1" then%>checked<%end if%>>��&nbsp;
</td>
<input type="hidden" id="b4" value="<b>�Ƿ���֤��̳</b><br><li>��֤��ֻ̳�й���Ա�͸ð�������ɼ��ͽ���<br><li>��֤��̳����֤�û�����Ӻ͹����ڰ��������������<br><li>�����˱�ѡ���ֻ����֤�û��ɽ���">
<td class=ForumrowHighlight><a href=# onclick="helpscript(b4);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>��������ƶ�</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(3)" value=0 <%if Board_Setting(3)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(3)" value=1 <%if Board_Setting(3)="1" then%>checked<%end if%>>����&nbsp;
</td>
<input type="hidden" id="b5" value="<b>��������ƶ�</b><br><li>����������Ա�Ϳ���Ȩ���û��ɽ����������<br><li>����������Ա�Ϳ���Ȩ���û���ֱ�ӷ���<br><li>һ���û�����˺����ӷ��ɼ�">
<td class=Forumrow><a href=# onclick="helpscript(b5);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>��չ����ƶ�</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(57)" value=0 <%if Board_Setting(57)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(57)" value=1 <%if Board_Setting(57)="1" then%>checked<%end if%>>����&nbsp;
<input type="hidden" id="bnew" value="<b>��չ��������ƶ�</b><br><li>����������Ա�Ϳ���Ȩ���û��ɽ����������<br><li>����������Ա�Ϳ���Ȩ���û���ֱ�ӷ���<br><li>һ���û��緢����������б����˵�����������˺����ӷ��ɼ�,<br>����ޱ����˵����ݣ��������˷�����">
</td>
<td class=Forumrow><a href=# onclick="helpscript(bnew);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>����������</U></td>
<td colspan=2 class=Forumrow>
<input type="text" Name=Board_Setting(58) Value="<%=Board_Setting(58)%>" Size=50><br>�����ö���������м���"|"�ָ��粻��д������0
<input type="hidden" id="bnewS" value="<b>����������</b><br><li>�����ö���������м��� | �ָ�">
</td>
<td class=Forumrow><a href=# onclick="helpscript(bnewS);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>����ͬʱ������</U><BR>������������Ϊ0</td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(18)" value="<%=Board_Setting(18)%>"> ��
</td>
<input type="hidden" id="b9" value="<b>����ͬʱ������</b><br><li>������������Ϊ0��������������ͬʱ������������̳�����������������ֵ�ʱ��δ��¼�û������ܷ��ʸð���">
<td class=ForumrowHighlight><a href=# onclick="helpscript(b9);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>��̳��ʱ���ã�</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(21)" value="0" <%If Board_Setting(21)="0" Then %>checked <%End If%>>�� ��</option>
<input type=radio name="Board_Setting(21)" value="1" <%If Board_Setting(21)="1" Then %>checked <%End If%>>��ʱ�ر�</option>
<input type=radio name="Board_Setting(21)" value="2" <%If Board_Setting(21)="2" Then %>checked <%End If%>>��ʱֻ��</option>
</td>
<input type="hidden" id="b10" value="<b>��ʱ����ѡ��:</b><br><li>�����������������Ƿ����ö�ʱ�ĸ��ֹ��ܣ���������˱����ܣ������ú�����ѡ���е���̳����ʱ�䣬��̳�ð��潫�����涨��ʱ������ָ��������">
<td class=Forumrow><a href=# onclick="helpscript(b10);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>��ʱ����</U><BR>�������Ҫѡ�񿪻��</td></td>
<td colspan=2 class=ForumrowHighlight>
<%
Board_Setting(22)=split(Board_Setting(22),"|")
If UBound(Board_Setting(22))<2 Then 
	Board_Setting(22)="1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1|1"
	Board_Setting(22)=split(Board_Setting(22),"|")
End If
For i= 0 to UBound(Board_Setting(22))
If i<10 Then Response.Write "&nbsp;"
%>
 <%=i%>�㣺<input type="checkbox" name="Board_Setting(22)<%=i%>" value="1" <%If Board_Setting(22)(i)="1" Then %>checked<%End If%>>��
   
 <%
 If (i+1) mod 4 = 0 Then Response.Write "<br>"
 Next
 %>
</td>
<input type="hidden" id="b11" value="<b>��̳����ʱ��</b><br><li>�����˱�ѡ�����ͬʱ���Ƿ����ö�ʱ������̳���ò���Ч�������˴�ѡ���̳�ð��潫�����涨��ʱ���ڸ��û�����">
<td class=ForumrowHighlight><a href=# onclick="helpscript(b11);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<%
Dim VisitConfirm
VisitConfirm=Split(Board_Setting(54),"|")
IF Ubound(VisitConfirm)<>8 Then
	Redim VisitConfirm(8)
	For i=0 To 8
	VisitConfirm(i)=0
	Next
End If
%>
<tr> 
<td colspan=2 class=Forumrow>
<U>�û�����������</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(0)%>">
</td>
<input type="hidden" id="VisitConfirm1" value="<b>�û�����������</b><br><li>���û���������´ﵽ������ʱ������ӵ�з���Ȩ�ޣ�<li>����������Ϊ0">
<td class=Forumrow><a href=# onclick="helpscript(VisitConfirm1);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�û����ٻ���</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(1)%>">
</td>
<input type="hidden" id="VisitConfirm2" value="<b>�û����ٻ���ֵ</b><br><li>���û��Ļ���ֵ�ﵽ������ʱ������ӵ�з���Ȩ�ޣ�<li>����������Ϊ0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(VisitConfirm2);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�û����ٽ�Ǯ</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(2)%>">
</td>
<input type="hidden" id="VisitConfirm3" value="<b>�û����ٽ�Ǯ��</b><br><li>���û��Ľ�Ǯ�ﵽ������ʱ������ӵ�з���Ȩ�ޣ�<li>����������Ϊ0">
<td class=Forumrow><a href=# onclick="helpscript(VisitConfirm3);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�û���������</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(3)%>">
</td>
<input type="hidden" id="VisitConfirm4" value="<b>�û���������</b><br><li>���û�������ֵ�ﵽ������ʱ������ӵ�з���Ȩ�ޣ�<li>����������Ϊ0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(VisitConfirm4);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�û���������</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(4)%>">
</td>
<input type="hidden" id="VisitConfirm5" value="<b>�û���������</b><br><li>���û������ﵽ������ʱ������ӵ�з���Ȩ�ޣ�<li>����������Ϊ0">
<td class=Forumrow><a href=# onclick="helpscript(VisitConfirm5);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�û����پ�������</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(5)%>">
</td>
<input type="hidden" id="VisitConfirm6" value="<b>�û����پ���������</b><br><li>���û�����ľ������´ﵽ������ʱ������ӵ�з���Ȩ�ޣ�<li>����������Ϊ0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(VisitConfirm6);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�û���ɾ����������</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(6)%>">
</td>
<input type="hidden" id="VisitConfirm7" value="<b>�û���ɾ����������</b><br><li>���û���ɾ����������������ʱ�����ܷ��ʸ÷ְ棡<li>����������Ϊ0">
<td class=Forumrow><a href=# onclick="helpscript(VisitConfirm7);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>����ע��ʱ�䣨��λΪ���ӣ�</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(7)%>">
</td>
<input type="hidden" id="VisitConfirm8" value="<b>�û�����ע��ʱ��</b><br><li>ע��ʱ����ָ�û�ע����ٷ��Ӻ�ɽ�����̳��<li>��λΪ���ӡ�<li>����������Ϊ0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(VisitConfirm8);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�����ϴ��ļ�����</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(54)" value="<%=VisitConfirm(8)%>">
</td>
<input type="hidden" id="VisitConfirm9" value="<b>�û������ϴ��ļ�����</b><br><li>���û������ϴ��ļ������ﵽ������ʱ������ӵ�з���Ȩ�ޣ�<li>����������Ϊ0">
<td class=Forumrow><a href=# onclick="helpscript(VisitConfirm9);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr><th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting3">ǰ̨����Ȩ��</a>[<a href="#top">����</a>]</th></tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>������������ɾ������</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(33)" value=0 <%if Board_Setting(33)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(33)" value=1 <%if Board_Setting(33)="1" then%>checked<%end if%>>��&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�����������޸Ĺ������</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(34)" value=0 <%if Board_Setting(34)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(34)" value=1 <%if Board_Setting(34)="1" then%>checked<%end if%>>��&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>���а��������޸Ĺ������</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(35)" value=0 <%if Board_Setting(35)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(35)" value=1 <%if Board_Setting(35)="1" then%>checked<%end if%>>��&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting4">�������</a>[<a href="#top">����</a>]</th>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�����Ƿ������֤��</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(4)" value=1 <%if Board_Setting(4)="1" then%>checked<%end if%>>����&nbsp;
<input type=radio name="Board_Setting(4)" value=0 <%if Board_Setting(4)="0" then%>checked<%end if%>>������&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�������Ƴ���</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(45)" value="<%=Board_Setting(45)%>"> Byte
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�����󷵻�</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(17)" value=1 <%if Board_Setting(17)="1" then%>checked<%end if%>>��ҳ&nbsp;
<input type=radio name="Board_Setting(17)" value=2 <%if Board_Setting(17)="2" then%>checked<%end if%>>��̳&nbsp;
<input type=radio name="Board_Setting(17)" value=3 <%if Board_Setting(17)="3" then%>checked<%end if%>>����&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>������������ֽ���</U><BR>1024�ֽڵ���1K</td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(16)" value="<%=Board_Setting(16)%>"> �ֽ�
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>����������С�ֽ���</U><BR>1024�ֽڵ���1K</td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(52)" value="<%=Board_Setting(52)%>"> �ֽ�
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>ͶƱ���Ƿ�ͶƱ�������������б���</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(53)" value=0 <%if Board_Setting(53)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(53)" value=1 <%if Board_Setting(53)="1" then%>checked<%end if%>>��&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�ϴ��ļ�����</U><BR>ÿ���ļ������á�|���ŷֿ�</td>
<td colspan=2 class=Forumrow>
<input type=text size=50 name="Board_Setting(19)" value="<%=Board_Setting(19)%>">
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�Ƿ����÷���ˮ����</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(30)" value=0 <%if Board_Setting(30)="0" then%>checked<%end if%>>��&nbsp;
<input type=radio name="Board_Setting(30)" value=1 <%if Board_Setting(30)="1" then%>checked<%end if%>>��&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>ÿ�η������</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(31)" value="<%=Board_Setting(31)%>"> ��
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>���ͶƱ��Ŀ</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(32)" value="<%=Board_Setting(32)%>"> ��
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr><th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting5">�����б���ʾ���</a>[<a href="#top">����</a>]</th></tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�����б������ʾ�ַ���</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(25)" value="<%=Board_Setting(25)%>">
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�����б�ÿҳ��¼��</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(26)" value="<%=Board_Setting(26)%>">
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�������ÿҳ��¼��</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(27)" value="<%=Board_Setting(27)%>">
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�����б�Ĭ�϶�ȡ������</U></td>
<td colspan=2 class=Forumrow>
<select size="1" name="Board_Setting(37)">
<option value="1"<%if Board_Setting(37)="0" then%> selected<%end if%>>ȫ����ʾ����</option>
<option value="2"<%if Board_Setting(37)="5" then%> selected<%end if%>>����������</option>
<option value="3"<%if Board_Setting(37)="15" then%> selected<%end if%>>����������</option>
<option value="4"<%if Board_Setting(37)="30" then%> selected<%end if%>>һ��������</option>
<option value="5"<%if Board_Setting(37)="60" then%> selected<%end if%>>����������</option>
<option value="6"<%if Board_Setting(37)="120" then%> selected<%end if%>>����������</option>
<option value="7"<%if Board_Setting(37)="180" then%> selected<%end if%>>����������</option>
</select>
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>��ʾ������ͼƬ��ʾ��ʽ</U></td>
<td colspan=2 class=ForumrowHighlight>
<select size="1" name="Board_Setting(38)">
<option value="0"<%if Board_Setting(38)="0" then%> selected<%end if%>>���ظ�ʱ��</option>
<option value="1"<%if Board_Setting(38)="1" then%> selected<%end if%>>����ʱ��</option>
</select>
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>��ʾ������ͼƬ��ʶʱ������</U></td>
<td colspan=2 class=Forumrow>
<select size="1" name="Board_Setting(61)">
<option value="0"<%if Board_Setting(61)="0" then%> selected<%end if%>>0����</option>
<option value="10"<%if Board_Setting(61)="10" then%> selected<%end if%>>10����</option>
<option value="30"<%if Board_Setting(61)="30" then%> selected<%end if%>>30����</option>
<option value="60"<%if Board_Setting(61)="60" then%> selected<%end if%>>1Сʱ</option>
<option value="360"<%If Board_Setting(61)="360" then%> selected<%end if%>>6Сʱ</option>
<option value="720"<%if Board_Setting(61)="720" then%> selected<%end if%>>12Сʱ</option>
<option value="1440"<%if Board_Setting(61)="1440" then%> selected<%end if%>>1��</option>
<option value="2880"<%if Board_Setting(61)="2880" then%> selected<%end if%>>2��</option>
</select>���ڸ��µ�����
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>��ʾ������ͼƬ��ַ���ã�new��:ֵΪ0���ʱ������ʾ����д׼ȷ��ַ��</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=30 name="Board_Setting(60)" value="<%=Board_Setting(60)%>">
<%
If instr(Board_Setting(60),".gif") Then Response.Write "<img src="""&Board_Setting(60)&""" border=0>"
%>
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting6">����������ʾ���</a>[<a href="#top">����</a>]</th>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>HTML�������</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(5)" value=0 <%if Board_Setting(5)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(5)" value=1 <%if Board_Setting(5)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>UBB�������</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(6)" value=0 <%if Board_Setting(6)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(6)" value=1 <%if Board_Setting(6)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>��ͼ��ǩ</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(7)" value=0 <%if Board_Setting(7)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(7)" value=1 <%if Board_Setting(7)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�����ǩ</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(8)" value=0 <%if Board_Setting(8)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(8)" value=1 <%if Board_Setting(8)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>Flash��ǩ</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(44)" value=0 <%if Board_Setting(44)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(44)" value=1 <%if Board_Setting(44)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>��ý���ǩ</U><BR>����RM,AVI��</td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(9)" value=0 <%if Board_Setting(9)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(9)" value=1 <%if Board_Setting(9)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�Ƿ񿪷Ž�Ǯ��</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(10)" value=0 <%if Board_Setting(10)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(10)" value=1 <%if Board_Setting(10)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�Ƿ񿪷Ż�����</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(11)" value=0 <%if Board_Setting(11)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(11)" value=1 <%if Board_Setting(11)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�Ƿ񿪷�������</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(12)" value=0 <%If Board_Setting(12)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(12)" value=1 <%If Board_Setting(12)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�Ƿ񿪷�������</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(13)" value=0 <%if Board_Setting(13)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(13)" value=1 <%if Board_Setting(13)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�Ƿ񿪷�������</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(14)" value=0 <%if Board_Setting(14)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(14)" value=1 <%if Board_Setting(14)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�Ƿ񿪷Żظ��ɼ���</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=radio name="Board_Setting(15)" value=0 <%if Board_Setting(15)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(15)" value=1 <%if Board_Setting(15)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�Ƿ񿪷ų������ӹ���</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(23)" value=0 <%if Board_Setting(23)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(23)" value=1 <%if Board_Setting(23)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�Ƿ񿪷Ŷ�Ա���ӹ���</U></td>
<td colspan=2 class=Forumrow>
<input type=radio name="Board_Setting(56)" value=0 <%if Board_Setting(56)="0" then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Board_Setting(56)" value=1 <%if Board_Setting(56)="1" then%>checked<%end if%>>����&nbsp;
</td>
<td class=Forumrow></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>���������ֺ�</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(28)" value="<%=Board_Setting(28)%>">
</td>
<td class=ForumrowHighlight></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>���������м��</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(29)" value="<%=Board_Setting(29)%>">
</td>
<td class=Forumrow></td>
</tr>
<%
Dim DownConfirm
DownConfirm=Split(Board_Setting(55),"|")
IF Ubound(DownConfirm)<>8 Then
	Redim DownConfirm(8)
	For i=0 To 8
	DownConfirm(i)=0
	Next
End If
%>
<tr><th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting7">���ظ�����������</a>[<a href="#top">����</a>]</th></tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�û�����������</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(0)%>">
</td>
<input type="hidden" id="Down1" value="<b>�û�����������</b><br><li>���û���������´ﵽ������ʱ������ӵ������Ȩ�ޣ�<li>����������Ϊ0">
<td class=Forumrow><a href=# onclick="helpscript(Down1);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�û����ٻ���</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(1)%>">
</td>
<input type="hidden" id="Down2" value="<b>�û����ٻ���ֵ</b><br><li>���û��Ļ���ֵ�ﵽ������ʱ������ӵ������Ȩ�ޣ�<li>����������Ϊ0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(Down2);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�û����ٽ�Ǯ</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(2)%>">
</td>
<input type="hidden" id="Down3" value="<b>�û����ٽ�Ǯ��</b><br><li>���û��Ľ�Ǯ�ﵽ������ʱ������ӵ������Ȩ�ޣ�<li>����������Ϊ0">
<td class=Forumrow><a href=# onclick="helpscript(Down3);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�û���������</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(3)%>">
</td>
<input type="hidden" id="Down4" value="<b>�û���������</b><br><li>���û�������ֵ�ﵽ������ʱ������ӵ������Ȩ�ޣ�<li>����������Ϊ0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(Down4);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�û���������</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(4)%>">
</td>
<input type="hidden" id="Down5" value="<b>�û���������</b><br><li>���û������ﵽ������ʱ������ӵ������Ȩ�ޣ�<li>����������Ϊ0">
<td class=Forumrow><a href=# onclick="helpscript(Down5);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>�û����پ�������</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(5)%>">
</td>
<input type="hidden" id="Down6" value="<b>�û����پ���������</b><br><li>���û�����ľ������´ﵽ������ʱ������ӵ������Ȩ�ޣ�<li>����������Ϊ0">
<td class=ForumrowHighlight><a href=# onclick="helpscript(Down6);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�û���ɾ����������</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(6)%>">
</td>
<input type="hidden" id="Down7" value="<b>�û���ɾ����������</b><br><li>���û���ɾ����������������ʱ���������ظð渽����<li>����������Ϊ0">
<td class=Forumrow><a href=# onclick="helpscript(Down7);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=ForumrowHighlight>
<U>����ע��ʱ��</U></td>
<td colspan=2 class=ForumrowHighlight>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(7)%>">
</td>
<input type="hidden" id="Down8" value="<b>�û�����ע������</b><br><li>���û�����ע����Ӵﵽ������ʱ������ӵ������Ȩ�ޣ�<li>�Է���Ϊ��λ��������Ϊ0��">
<td class=ForumrowHighlight><a href=# onclick="helpscript(Down8);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�����ϴ��ļ�����</U></td>
<td colspan=2 class=Forumrow>
<input type=text size=10 name="Board_Setting(55)" value="<%=DownConfirm(8)%>">
</td>
<input type="hidden" id="Down9" value="<b>�û������ϴ��ļ�����</b><br><li>���û������ϴ��ļ������ﵽ������ʱ������ӵ������Ȩ�ޣ�<li>����������Ϊ0">
<td class=Forumrow><a href=# onclick="helpscript(Down9);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr><th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting9">��̳������������</a>[<a href="#top">����</a>]</th></tr>
<tr> 
<td colspan=2 class=Forumrow>
<U>�����������������</U></td>
<td colspan=2 class=Forumrow>
<input type="radio" name="Board_Setting(59)" value="0"
<%
If Board_Setting(59)="0" Then
%>
 checked 
 <%
 End If
 %>
 >&nbsp;��ʾȫ��
<input type="radio" name="Board_Setting(59)" value="1"
<%
If Board_Setting(59)="1" Then
%>
 checked 
 <%
 End If
 %>
 >&nbsp;��ʾ��
 <input type="radio" name="Board_Setting(59)" value="2"
<%
If Board_Setting(59)="2" Then
%>
 checked 
 <%
 End If
 %>
 >&nbsp;��ʾ����
 <input type="radio" name="Board_Setting(59)" value="3"
<%
If Board_Setting(59)="3" Then
%>
 checked 
 <%
 End If
 %>
 >&nbsp;����ʾ����ʾͷ��
</td>
<input type="hidden" id="xx9" value="<b>�û������ϴ��ļ�����</b><br><li>���û������ϴ��ļ������ﵽ������ʱ������ӵ������Ȩ�ޣ�<li>����������Ϊ0">
<td class=Forumrow><a href=# onclick="helpscript(xx9);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr><th height="25" colspan="5" id=tabletitlelink align=left>  &nbsp;<a name="setting8">��̳ר������������</a>[<a href="#top">����</a>]</th></tr>
<tr><td colspan="5" class=Forumrow>
<li>������ר��Ȩ�ޣ��뵽��Ӧ�û��鷢��Ȩ�������ã�
<li>ר����Ŀ������ӣ��޸ģ�
<li>ע��ɾ��ר��ͬʱ,�Ὣ���ר����������¸���Ϊ��ͨ���⡣</td></tr>
<%
Dim BoardTopic,BoardTopicImg,ii
BoardTopic=Split(Board_Setting(48),"$$")
BoardTopicImg=Split(Board_Setting(49),"$$")
For ii=0 to Ubound(BoardTopic)-1
%>
<tr>
<td width="15%" class=Forumrow><U>ר������:</U></td>
<td width="35%" class=Forumrow>
<input type=text size=30 name="BoardTopic" value="<%=Server.Htmlencode(BoardTopic(ii))%>"></td>
<td width="15%" class=Forumrow><U>��Ӧ��ʾͼ�꣺</U></td>
<td width="35%" class=Forumrow>
<input type=text size=30 name="BoardTopicImg" value="<%=BoardTopicImg(ii)%>">
<%
If BoardTopicImg(ii)<>"" and Instr(BoardTopicImg(ii),".gif") Then Response.Write "<img src="&BoardTopicImg(ii)&" border=0>"
%>
</td>
<td class=Forumrow></td>
</tr>
<%Next%>
<input type=hidden value="<%=ii%>" name="BoardTopicNum">
<tr>
<td width="15%" class=Forumrow><U>���ר��:</U></td>
<td width="35%" class=Forumrow>
<input type=text size=30 name="BoardTopic" value=""></td>
<td width="15%" class=Forumrow><U>��Ӧ��ʾͼ�꣺</U></td>
<td width="35%" class=Forumrow>
<input type=text size=30 name="BoardTopicImg" value=""></td>
<td class=Forumrow></td>
</tr>
<tr>
<td colspan=5 class=ForumRowHighlight>
<div align="center"> 
<input type=hidden value="<%=Board_Setting(20)%>" name="Board_Setting(20)">
<input type=hidden value="<%=Board_Setting(46)%>" name="Board_Setting(46)">
<input type=hidden value="<%=Board_Setting(47)%>" name="Board_Setting(47)">
<input type="submit" name="Submit" value="�� ��">
</div>
</td>
</tr>
</form>
</table>
<%
end sub

sub saveconst()
if not isnumeric(request("editid")) then
Errmsg=ErrMsg + "<BR><li>����İ������"
dvbbs_error()
exit sub
else
Dim iboard_setting,isetting
Dim BoardTopic,BoardTopicImg,TempStr,ii,BoardTopicNum
Dim DownConfirm,ViewConfirm
	ii=0
	i=0
	For Each TempStr in Request.Form("Board_Setting(54)")
		i=i+1
		ViewConfirm=ViewConfirm&TempStr
		If i<>Request.Form("Board_Setting(54)").count Then
		ViewConfirm=ViewConfirm&"|"
		End If
	Next
	i=0
	If not ISNumeric(Replace(ViewConfirm,"|","")) or Request.Form("Board_Setting(54)").count<>9 Then
		Errmsg=ErrMsg + "<BR><li>���ظ��������д����ύ����ֹ��"
		dvbbs_error()
		exit sub
	End if
	For Each TempStr in Request.Form("Board_Setting(55)")
		i=i+1
		DownConFirm=DownConFirm&TempStr
		If i<>Request.Form("Board_Setting(55)").count Then
		DownConFirm=DownConFirm&"|"
		End If
	Next
	i=0
	If not ISNumeric(Replace(DownConFirm,"|","")) or Request.Form("Board_Setting(55)").count<>9 Then
		Errmsg=ErrMsg + "<BR><li>���ظ��������д����ύ����ֹ��"
		dvbbs_error()
		exit sub
	End if
	
	IF Request("BoardTopicNum")<>"" and Isnumeric(Request("BoardTopicNum")) Then
	BoardTopicNum=Request("BoardTopicNum") 
	Else
	BoardTopicNum=0
	End If
	For Each TempStr in Request.form("BoardTopic")
		If TempStr<>"" Then 
			BoardTopic=BoardTopic&TempStr&"$$"
			ii=ii+1
		End If
	Next
	TempStr=""
	For Each TempStr in Request.form("BoardTopicImg")
			BoardTopicImg=BoardTopicImg&TempStr&"$$"
	Next
	TempStr=""
	If ii>99 Then
		Errmsg=ErrMsg + "<BR><li>ר����Ŀ��Ŀ�ڣ��������ڡ�"
		dvbbs_error()
		exit sub
	End If
	Dim setingdata,j
	For i = 0 To 70
		If Trim(request.Form("Board_Setting("&i&")"))="" Or i=22 Then
			'Response.Write "Board_Setting("&i&")<br>"
			isetting=0
			If i=22 Then
				isetting=""
				For j=0 to  23
					If isetting="" Then
						If Request.form("Board_Setting(22)"&j)="1" Then
							isetting="1"
						Else
							isetting="0"
						End If
					Else
						If Request.form("Board_Setting(22)"&j)="1" Then
							isetting=isetting&"|1"
						Else
							isetting=isetting&"|0"
						End If
					End If
				Next
			End If
		Else
			isetting=Replace(Trim(request.Form("Board_Setting("&i&")")),",","")
		End If
		If i = 0 Then
			iboard_Setting = isetting
		ElseIf i = 48 Then
			iboard_Setting = iboard_Setting & "," & BoardTopic
		ElseIf i = 49 Then
			iboard_Setting = iboard_Setting & "," & BoardTopicImg
		ElseIf i=54 Then
			iboard_Setting = iboard_Setting & "," & ViewConfirm
		ElseIf i=55 Then 
			iboard_Setting = iboard_Setting & "," & DownConFirm
		Else
			iboard_Setting = iboard_Setting & "," & isetting
		End If
	Next

Dim FoundCKBoard
FoundCKBoard=False
For i=0 to UBOUND(Dvbbs.Forum_Setting)
	If request.Form("CK_Board_Setting("&i&")")<>"" Then
		FoundCKBoard=True
		Exit For
	End If
Next

Dim Forum_Boards,upBoardid,upid,temprs
select case request("savetype")
'��ǰ����
case "0"
	Dvbbs.Execute("update dv_board set board_setting='"&iboard_setting&"' where boardid="&Request("editid"))
	Dvbbs.ReloadBoardInfo(Request("editid"))
	upBoardid=" and boardid="&Request("editid")
'���а���
case "1"
	Dvbbs.Execute("update dv_board set board_setting='"&iboard_setting&"'")
	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.ReloadBoardInfo(Forum_Boards(i))
	Next
	upBoardid=""
'��ͬ���������а��棨���������ࣩ
case "2"
	set rs=Dvbbs.Execute("select rootid from dv_board where boardid="&request("editid"))
	if not rs.eof then
		Dvbbs.Execute("update dv_board set board_setting='"&iboard_setting&"' where (Not ParentID=0) and rootid="&rs(0))
		Set temprs=Dvbbs.Execute("Select boardid from Dv_board where (Not ParentID=0) and rootid="&rs(0))
		if not temprs.eof then
			upid=temprs.GetString(,, "",",","")
		end if
		temprs.close:Set temprs=Nothing
	end if
	rs.close:set rs=nothing
	upBoardid=" and boardid in ("&left(upid,(len(upid)-1))&")"
	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.ReloadBoardInfo(Forum_Boards(i))
	Next
'��ͬ���������а��棨�������ࣩ
case "3"
	set rs=Dvbbs.Execute("select rootid from dv_board where boardid="&request("editid"))
	if not rs.eof then
		Dvbbs.Execute("update dv_board set board_setting='"&iboard_setting&"' where rootid="&rs(0))
		Set temprs=Dvbbs.Execute("select boardid from dv_board where rootid="&rs(0))
		if not temprs.eof then
			upid=temprs.GetString(,, "",",","")
		end if
		temprs.close:Set temprs=Nothing
	end if
	rs.close:set rs=nothing
	upBoardid=" and boardid in ("&left(upid,(len(upid)-1))&")"

	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.ReloadBoardInfo(Forum_Boards(i))
	Next
'ͬ����ͬ�������
case "4"
	set rs=Dvbbs.Execute("select rootid,ParentStr,ParentID from dv_board where boardid="&request("editid"))
	if not rs.eof then
		Dvbbs.Execute("update dv_board set board_setting='"&iboard_setting&"' where rootid="&rs(0)&" and ParentID="&rs(2)&" and ParentStr='"&rs(1)&"'")
		Set temprs=Dvbbs.Execute("select boardid from dv_board where rootid="&rs(0)&" and ParentID="&rs(2)&" and ParentStr='"&rs(1)&"'")
		if not temprs.eof then
			upid=temprs.GetString(,, "",",","")
		end if
		temprs.close:Set temprs=Nothing
	end if
	rs.close:set rs=nothing
	upBoardid=" and boardid in ("&left(upid,(len(upid)-1))&")"
	Forum_Boards=Split(Dvbbs.CacheData(27,0),",")
	For i=0 To Ubound(Forum_Boards)
		Dvbbs.ReloadBoardInfo(Forum_Boards(i))
	Next
End Select

If BoardTopicNum>ii Then
	Dvbbs.Execute("update Dv_Topic set Mode=0 where Mode >= "&ii+1&" "&upBoardid&" ")
End If

dv_suc("���óɹ���<a href=admin_boardsetting.asp?editid="&request("editid")&">���ذ���߼�����</a>")
End If
End sub
%>
