<%
Function GroupPermission(GroupSetting)
Dim reGroupSetting,Rs,UserHtml,UserHtmlA,UserHtmlB
If GroupSetting="" Then
	Set Rs = Dvbbs.Execute("Select GroupSetting From Dv_UserGroups Where UserGroupID=4")
	reGroupSetting = Split(Rs(0),",")
Else
	reGroupSetting = Split(GroupSetting,",")
End If
If reGroupSetting(58)="0" Then reGroupSetting(58)="��"
UserHtml = Split(reGroupSetting(58),"��")
If Ubound(UserHtml)=1 Then
	UserHtmlA=UserHtml(0)
	UserHtmlB=UserHtml(1)
Else
	UserHtmlA=""
	UserHtmlB=""
End If
%>
<tr> 
<th height="23" colspan="3"  align=left>����������ѡ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�û�����������������ʾ���<BR>HTML�﷨�����ұ�Ǵ��뽫�����û���ǰ����ͷ</td>
<td height="23" width="40%" class=Forumrow>���� <input name="GroupSetting(58)A" type=text size=8 value="<%=Server.HtmlEncode(UserHtmlA)%>"> �ұ�� <input name="GroupSetting(58)B" type=text size=8 value="<%=Server.HtmlEncode(UserHtmlB)%>"></td>
<input type="hidden" id="g1" value="<b>�û�����������������ʾ���</b><br><li>HTML�﷨�����ұ�Ǵ��뽫�����û���ǰ����ͷ<br><li>����������ǰ��ֱ�Ϊ��b���͡�/b�����������������и����û�������صȼ��û�����ʾΪ<B>����</B>">
<td class=forumRow><a href=# onclick="helpscript(g1);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>�����û���ѡ���</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(57)" type=radio value="1" <%if reGroupSetting(57)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(57)" type=radio value="0" <%if reGroupSetting(57)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g2" value="<b>�����û���ѡ���</b><br><li>����ر��˱�ѡ���̳���û��������Լ�ѡ�������ʾ�ķ�񣨰����û��ڸ�����Ϣ���趨�ķ��">
<td class=forumRowHighlight><a href=# onclick="helpscript(g2);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���������̳</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(0)" type=radio value="1" <%if reGroupSetting(0)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(0)" type=radio value="0" <%if reGroupSetting(0)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g3" value="<b>�û�����������������ʾ���</b><br><li>�رմ�ѡ�������ȼ��û������������̳<br><li>ʹ�ü��ɣ��������趨ĳ���û��鲻��ʹ�ñ����ã���������ݱ仯����û����ʹ�ñ����ã������ÿ��˲���ʹ�ñ����ã���������ʹ����¼">
<td class=forumRow><a href=# onclick="helpscript(g3);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Բ鿴��Ա��Ϣ(����������Ա�����Ϻͻ�Ա�б�)
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(1)" type=radio value="1" <%if reGroupSetting(1)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(1)" type=radio value="0" <%if reGroupSetting(1)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g4" value="<b>���Բ鿴��Ա��Ϣ</b><br><li>�رմ�ѡ�������ȼ��û������������̳�û����ϣ�������Ա���Ϻͻ�Ա�б�����<br><li>ʹ�ü��ɣ��������趨ĳ���û��鲻��ʹ�ñ����ã���������ݱ仯����û����ʹ�ñ����ã������ÿ��˲���ʹ�ñ����ã���������ʹ����¼">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g4);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Բ鿴�����˷���������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(2)" type=radio value="1" <%if reGroupSetting(2)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(2)" type=radio value="0" <%if reGroupSetting(2)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g5" value="<b>���Բ鿴�����˷���������</b><br><li>�رմ�ѡ�������ȼ��û������������̳�������˷���������<br><li>ʹ�ü��ɣ��������趨ĳ���û��鲻��ʹ�ñ����ã���������ݱ仯����û����ʹ�ñ����ã������ÿ��˲���ʹ�ñ����ã���������ʹ����¼">
<td class=Forumrow><a href=# onclick="helpscript(g5);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���������������
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(41)" type=radio value="1" <%if reGroupSetting(41)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(41)" type=radio value="0" <%if reGroupSetting(41)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g6" value="<b>���������������</b><br><li>�رմ�ѡ�������ȼ��û������������̳�еľ�������<br><li>ʹ�ü��ɣ��������趨ĳ���û��鲻��ʹ�ñ����ã���������ݱ仯����û����ʹ�ñ����ã������ÿ��˲���ʹ�ñ����ã���������ʹ����¼">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g6);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<th height="23" colspan="3"  align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է���������</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(3)" type=radio value="1" <%if reGroupSetting(3)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(3)" type=radio value="0" <%if reGroupSetting(3)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g9" value="<b>���Է���������</b><br><li>�򿪴�ѡ�������ȼ��û������Կ��Է��������⡣���ڹ��ҹ涨����̳Ĭ�ϵ�δ��¼�û��齫��ʹ���ô�ѡ��Ҳ���ܷ���">
<td class=Forumrow><a href=# onclick="helpscript(g9);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>�����ģʽ�¿�ֱ�ӷ��������辭�����</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(64)" type=radio value="1" <%if reGroupSetting(64)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(64)" type=radio value="0" <%if reGroupSetting(64)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g10" value="<b>�����ģʽ�¿�ֱ�ӷ��������辭�����</b><br><li>�򿪴�ѡ�������ȼ��û������Կ��Է����������ظ����������<br><li>����̳��������Ϊ���״̬ʱ��ѡ����Ч">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g10);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>һ����෢����Ŀ
</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(62)" type=text size=4 value="<%=reGroupSetting(62)%>"></td>
<input type="hidden" id="g11" value="<b>һ����෢����Ŀ</b><br><li>��д0Ϊ�������ƣ����ڶԸ���ˮ����ʹ������������û������ڴ����ú��������<br><li>ʹ�ü��ɣ������Ը���ͬ�û������ò�ͬ������">
<td class=Forumrow><a href=# onclick="helpscript(g11);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Իظ��Լ�������
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(4)" type=radio value="1" <%if reGroupSetting(4)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(4)" type=radio value="0" <%if reGroupSetting(4)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g12" value="<b>���Իظ��Լ�������</b><br><li>�򿪴�ѡ�����û����ȼ��û����Իظ��Լ�����������">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g12);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Իظ������˵�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(5)" type=radio value="1" <%if reGroupSetting(5)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(5)" type=radio value="0" <%if reGroupSetting(5)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g13" value="<b>���Իظ������˵�����</b><br><li>�򿪴�ѡ�����û����ȼ��û����Իظ������˵�����">
<td class=Forumrow><a href=# onclick="helpscript(g13);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>��������̳�������ֵ�ʱ���������(�ʻ��ͼ���)?
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(6)" type=radio value="1" <%if reGroupSetting(6)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(6)" type=radio value="0" <%if reGroupSetting(6)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g14" value="<b>��������̳�������ֵ�ʱ���������</b><br><li>�򿪴�ѡ�����û����ȼ��û���������̳�������ֵ�ʱ��������֣�Ҳ�������������е��ʻ��ͼ���ѡ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g14);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�������������Ǯ
</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(47)" type=text size=4 value="<%=reGroupSetting(47)%>"></td>
<input type="hidden" id="g15" value="<b>�������������Ǯ</b><br><li>����û����ȼ��û��������������Ǯ������ʻ��򼦵���۳�">
<td class=Forumrow><a href=# onclick="helpscript(g15);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Է�����ͶƱ</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(8)" type=radio value="1" <%if reGroupSetting(8)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(8)" type=radio value="0" <%if reGroupSetting(8)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g21" value="<b>���Է�����ͶƱ</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Է�����ͶƱ">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g21);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Բ���ͶƱ</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(9)" type=radio value="1" <%if reGroupSetting(9)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(9)" type=radio value="0" <%if reGroupSetting(9)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g22" value="<b>���Է�����ͶƱ</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Բ���ͶƱ">
<td class=Forumrow><a href=# onclick="helpscript(g22);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>ͶƱ����ʹ��HTML�﷨</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(68)" type=radio value="1" <%if reGroupSetting(68)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(68)" type=radio value="0" <%if reGroupSetting(68)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g_u_HTML" value="<b>ͶƱ����ʹ��HTML�﷨</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ������ͶƱ��ʹ��HTML�﷨">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g_u_HTML);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Է���С�ֱ�</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(17)" type=radio value="1"  <%if reGroupSetting(17)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(17)" type=radio value="0" <%if reGroupSetting(17)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g23" value="<b>���Է���С�ֱ�</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Է���С�ֱ�">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g23);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����С�ֱ������Ǯ</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(46)" type=text value="<%=reGroupSetting(46)%>" size=4></td>
<input type="hidden" id="g24" value="<b>����С�ֱ������Ǯ</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û�����С�ֱ������Ǯ">
<td class=Forumrow><a href=# onclick="helpscript(g24);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Է�������������ӣ������Ӻ졢UBB�﷨�ȣ�</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(51)" type=radio value="1"  <%if reGroupSetting(51)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(51)" type=radio value="0" <%if reGroupSetting(51)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g25" value="<b>���Է��������������</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û����Է�������������ӣ���������ɫ��HTML�﷨��UBB�﷨�ȣ�������Ը����û����ʹ�ô����⹦��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g25);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<!--
<tr>
<td height="23" width="60%" class=Forumrow>����ģʽѡ��</td>
<td height="23" width="40%" class=Forumrow>
<select name="GroupSetting(67)" >
<option value="0"  <%if reGroupSetting(67)="0" then%>selected<%end if%>>�ر�HTML�༭
<option value="1"  <%if reGroupSetting(67)="1" then%>selected<%end if%>>����HTML�༭
<option value="2"  <%if reGroupSetting(67)="2" then%>selected<%end if%>>��ģʽ�༭
<option value="3"  <%if reGroupSetting(67)="3" then%>selected<%end if%>>ȫ���ܱ༭
</select>
</td>
<input type="hidden" id="g0" value="<b>����ģʽѡ��</b><br><li>����ģʽ������Design�༭ģʽ,Ubb��ģʽ��HTML�ɱ༭ģʽ��<li>�ر�HTML�༭�������������߼�ģʽ�£��û�ֻ����Design�༭ģʽ��Ubb��ģʽ��<li>����HTML�༭�������������߼�ģʽ�£��û�ӵ��ȫ���ܱ༭ģʽ��<li>��ģʽ�༭�������������߼�ģʽ�£��û�ֻ����Ubb��ģʽ��<li>ȫ���ܱ༭��������ڷ����ģʽ�£�ӵ�����з���ģʽ��<li>Ϊ�����û�����HTML�ĸ����﷨������ֻ�Բ����û��ر�HTML�༭��">
<td class=Forumrow><a href=# onclick="helpscript(g0);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
-->
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Է�����̳ר��</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(65)" type=radio value="1"  <%if reGroupSetting(65)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(65)" type=radio value="0" <%if reGroupSetting(65)="0" then%>checked<%end if%>></td>
<td class=ForumrowHighLight><a href=# class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��ע���û����ٷ��Ӻ���ܷ���</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(52)" type=text value="<%=reGroupSetting(52)%>" size=4> ����</td>
<input type="hidden" id="g26" value="<b>��ע���û����ٷ��Ӻ���ܷ���</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û���ע����Ҫ���ٷ��Ӻ���ܷ��ԣ�����������ô�ѡ��Ա���һЩ�����û���ע��ɢ���Ƿ����ӻ�������">
<td class=Forumrow><a href=# onclick="helpscript(g26);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<th height="23" colspan="3"  align=left>����<b>����/����༭Ȩ��</b></th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Ա༭�Լ�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(10)" type=radio value="1" <%if reGroupSetting(10)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(10)" type=radio value="0" <%if reGroupSetting(10)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g27" value="<b>���Ա༭�Լ�������</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Ա༭�Լ�������">
<td class=Forumrow><a href=# onclick="helpscript(g27);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>����ɾ���Լ�������
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(11)" type=radio value="1" <%if reGroupSetting(11)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(11)" type=radio value="0" <%if reGroupSetting(11)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g28" value="<b>����ɾ���Լ�������</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ����ɾ���Լ������ӣ�������Լ�����Ҫ�������ô�ѡ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g28);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����ƶ��Լ������ӵ�������̳
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(12)" type=radio value="1" <%if reGroupSetting(12)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(12)" type=radio value="0" <%if reGroupSetting(12)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g29" value="<b>�����ƶ��Լ������ӵ�������̳</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ�����ƶ��Լ������ӵ�������̳��������Լ�����Ҫ�������ô�ѡ��">
<td class=Forumrow><a href=# onclick="helpscript(g29);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Դ�/�ر��Լ�����������
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(13)" type=radio value="1" <%if reGroupSetting(13)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(13)" type=radio value="0" <%if reGroupSetting(13)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g30" value="<b>���Դ�/�ر��Լ�����������</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Դ�/�ر��Լ����������⣬������Լ�����Ҫ�������ô�ѡ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g30);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<th height="23" colspan="3" align=left>�����ϴ�Ȩ������</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����ϴ�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(7)" type=radio value="1" <%if reGroupSetting(7)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(7)" type=radio value="0" <%if reGroupSetting(7)="0" then%>checked<%end if%>>
&nbsp;���������ϴ�<input name="GroupSetting(7)" type=radio value="2" <%if reGroupSetting(7)="2" then%>checked<%end if%>>&nbsp;�ظ������ϴ�<input name="GroupSetting(7)" type=radio value="3" <%if reGroupSetting(7)="3" then%>checked<%end if%>>
</td>
<input type="hidden" id="g16" value="<b>�����ϴ�����</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ�����ϴ�������ѡ���������ͻ����������ϴ��������С���Ҳ���Կ��Ը�����Ҫ�ֱ����÷���������Ƿ�����ϴ�">
<td class=Forumrow><a href=# onclick="helpscript(g16);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>һ�������ϴ�����(����Ϊ0����������ʹ�ô˹���;���鲻Ҫ����5��)
</td>
<td height="23" width="40%" class=ForumrowHighLight><input name="GroupSetting(66)" type=text size=4 value="<%=reGroupSetting(66)%>"></td>
<input type="hidden" id="GroupSetting66" value="<b>һ�������ϴ�����</b><br><li>����Ϊ0����������ʹ�ô˹���;<li>���鲻Ҫ����5������Ϊ�ϴ����������Ĵ�����������Դ">
<td class=ForumrowHighLight><a href=# onclick="helpscript(GroupSetting66);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>һ������ϴ��ļ�����
</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(40)" type=text size=4 value="<%=reGroupSetting(40)%>"></td>
<input type="hidden" id="g17" value="<b>һ������ϴ��ļ�����</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û�һ������ϴ��ļ����������鲻Ҫ���ù�����Ϊ�ϴ����������Ĵ�����������Դ">
<td class=Forumrow><a href=# onclick="helpscript(g17);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>һ������ϴ��ļ�����
</td>
<td height="23" width="40%" class=ForumrowHighLight><input name="GroupSetting(50)" type=text size=4 value="<%=reGroupSetting(50)%>"></td>
<input type="hidden" id="g18" value="<b>һ������ϴ��ļ�����</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û�һ������ϴ��ļ����������鲻Ҫ���ù�����Ϊ�ϴ����������Ĵ�����������Դ">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g18);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�ϴ��ļ���С����
</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(44)" type=text size=4 value="<%=reGroupSetting(44)%>"> KB</td>
<input type="hidden" id="g19" value="<b>�ϴ��ļ���С����</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��ϴ��ļ���С�����鲻Ҫ���ù�����Ϊ�ϴ����������Ĵ�����������Դ">
<td class=Forumrow><a href=# onclick="helpscript(g19);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>�������ظ���</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(61)" type=radio value="1" <%if reGroupSetting(61)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(61)" type=radio value="0" <%if reGroupSetting(61)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g20" value="<b>�������ظ���</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ�������ظ����������������δ��¼�û���������">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g20);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<th height="23" colspan="3" align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����ɾ������������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(18)" type=radio value="1" <%if reGroupSetting(18)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(18)" type=radio value="0"  <%if reGroupSetting(18)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g38" value="<b>����ɾ������������</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ����ɾ�����������ӣ�������Լ�����Ҫ�������ô�ѡ�����԰������������û������ô�Ȩ��">
<td class=Forumrow><a href=# onclick="helpscript(g38);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>�����ƶ�����������
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(19)" type=radio value="1" <%if reGroupSetting(19)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(19)" type=radio value="0"  <%if reGroupSetting(19)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g39" value="<b>�����ƶ�����������</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ�����ƶ����������ӣ�������Լ�����Ҫ�������ô�ѡ�����԰������������û������ô�Ȩ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g39);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Դ�/�ر�����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(20)" type=radio value="1" <%if reGroupSetting(20)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(20)" type=radio value="0"  <%if reGroupSetting(20)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g40" value="<b>���Դ�/�ر�����������</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Դ�/�ر����������ӣ�������Լ�����Ҫ�������ô�ѡ�����԰������������û������ô�Ȩ��">
<td class=Forumrow><a href=# onclick="helpscript(g40);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Թ̶�/����̶�����
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(21)" type=radio value="1" <%if reGroupSetting(21)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(21)" type=radio value="0"  <%if reGroupSetting(21)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g41" value="<b>���Թ̶�/����̶�����</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Թ̶�/����̶����ӣ�������Լ�����Ҫ�������ô�ѡ�����԰������������û������ô�Ȩ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g41);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Խ�����������̶�����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(54)" type=radio value="1" <%if reGroupSetting(54)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(54)" type=radio value="0"  <%if reGroupSetting(54)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g42" value="<b>���Խ�����������̶�����</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Խ�����������̶�������������Լ�����Ҫ�������ô�ѡ�����Գ����������������û������ô�Ȩ��">
<td class=Forumrow><a href=# onclick="helpscript(g42);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Խ��������̶ܹ�����
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(38)" type=radio value="1"  <%if reGroupSetting(38)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(38)" type=radio value="0" <%if reGroupSetting(38)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g43" value="<b>���Խ��������̶ܹ�����</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Խ��������̶ܹ�������������Լ�����Ҫ�������ô�ѡ�����Գ����������������û������ô�Ȩ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g43);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Խ���/�ͷ������û�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(22)" type=radio value="1" <%if reGroupSetting(22)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(22)" type=radio value="0"  <%if reGroupSetting(22)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g44" value="<b>���Խ���/�ͷ������û�</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Խ���/�ͷ������û���������Լ�����Ҫ�������ô�ѡ�����԰������������û������ô�Ȩ��">
<td class=Forumrow><a href=# onclick="helpscript(g44);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Խ���/�ͷ��û�
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(43)" type=radio value="1" <%if reGroupSetting(43)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(43)" type=radio value="0"  <%if reGroupSetting(43)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g45" value="<b>���Խ���/�ͷ��û�</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Խ���/�ͷ��û���������Լ�����Ҫ�������ô�ѡ�����Գ����������������û������ô�Ȩ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g45);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Ա༭����������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(23)" type=radio value="1" <%if reGroupSetting(23)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(23)" type=radio value="0" <%if reGroupSetting(23)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g46" value="<b>���Ա༭����������</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Ա༭���������ӣ�������Լ�����Ҫ�������ô�ѡ�����԰������������û������ô�Ȩ��">
<td class=Forumrow><a href=# onclick="helpscript(g46);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Լ���/�����������
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(24)" type=radio value="1" <%if reGroupSetting(24)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(24)" type=radio value="0"  <%if reGroupSetting(24)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g47" value="<b>���Լ���/�����������</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Լ���/����������ӣ�������Լ�����Ҫ�������ô�ѡ�����԰������������û������ô�Ȩ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g47);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(25)" type=radio value="1" <%if reGroupSetting(25)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(25)" type=radio value="0"  <%if reGroupSetting(25)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g48" value="<b>���Է�������</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Է������棬������Լ�����Ҫ�������ô�ѡ�����԰������������û������ô�Ȩ��">
<td class=Forumrow><a href=# onclick="helpscript(g48);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Թ�����
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(26)" type=radio value="1" <%if reGroupSetting(26)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(26)" type=radio value="0"  <%if reGroupSetting(26)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g49" value="<b>���Թ�����</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Թ����棬������Լ�����Ҫ�������ô�ѡ�����Գ����������������û������ô�Ȩ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g49);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Թ���С�ֱ�
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(27)" type=radio value="1" <%if reGroupSetting(27)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(27)" type=radio value="0"  <%if reGroupSetting(27)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g50" value="<b>���Թ���С�ֱ�</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Թ���С�ֱ���������Լ�����Ҫ�������ô�ѡ�����Գ����������������û������ô�Ȩ��">
<td class=Forumrow><a href=# onclick="helpscript(g50);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>��������/����/��������û�
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(28)" type=radio value="1" <%if reGroupSetting(28)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(28)" type=radio value="0"  <%if reGroupSetting(28)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g51" value="<b>��������/����/��������û�</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ��������/����/��������û���������Լ�����Ҫ�������ô�ѡ�����Գ����������������û������ô�Ȩ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g51);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����ɾ���û�1��10������������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(29)" type=radio value="1" <%if reGroupSetting(29)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(29)" type=radio value="0"  <%if reGroupSetting(29)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g52" value="<b>����ɾ���û�1��10������������</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ����ɾ���û�1��10�����������ӣ�������Լ�����Ҫ�������ô�ѡ�����Գ����������������û������ô�Ȩ��">
<td class=Forumrow><a href=# onclick="helpscript(g52);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Բ鿴����IP����Դ
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(30)" type=radio value="1" <%if reGroupSetting(30)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(30)" type=radio value="0"  <%if reGroupSetting(30)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g53" value="<b>���Բ鿴����IP����Դ</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Բ鿴����IP����Դ��������Լ�����Ҫ�������ô�ѡ�����԰������������û������ô�Ȩ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g53);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����޶�IP����
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(31)" type=radio value="1" <%if reGroupSetting(31)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(31)" type=radio value="0"  <%if reGroupSetting(31)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g54" value="<b>�����޶�IP����</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ�����޶�IP���ã�������Լ�����Ҫ�������ô�ѡ�����Գ����������������û������ô�Ȩ��">
<td class=Forumrow><a href=# onclick="helpscript(g54);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���Թ����û�Ȩ��
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(42)" type=radio value="1" <%if reGroupSetting(42)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(42)" type=radio value="0"  <%if reGroupSetting(42)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g55" value="<b>���Թ����û�Ȩ��</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���Թ����û�Ȩ�ޣ�������Լ�����Ҫ�������ô�ѡ�����Գ����������������û������ô�Ȩ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g55);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��������ɾ�����ӣ�ǰ̨��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(45)" type=radio value="1" <%if reGroupSetting(45)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(45)" type=radio value="0"  <%if reGroupSetting(45)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g56" value="<b>��������ɾ�����ӣ�ǰ̨��</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ��������ɾ�����ӣ�ǰ̨����������Լ�����Ҫ�������ô�ѡ�����԰������������û������ô�Ȩ��">
<td class=Forumrow><a href=# onclick="helpscript(g56);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>�Ƿ���������ӵ�Ȩ��
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(36)" type=radio value="1" <%if reGroupSetting(36)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(36)" type=radio value="0" <%if reGroupSetting(36)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g57" value="<b>�Ƿ���������ӵ�Ȩ��</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���������ӵ�Ȩ�ޣ�������Լ�����Ҫ�������ô�ѡ�����԰������������û������ô�Ȩ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g57);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�Ƿ��н���������̳��Ȩ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(37)" type=radio value="1"  <%if reGroupSetting(37)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(37)" type=radio value="0" <%if reGroupSetting(37)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g58" value="<b>�Ƿ��н���������̳��Ȩ��</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ��н���������̳��Ȩ�ޣ�������Լ�����Ҫ�������ô�ѡ��">
<td class=Forumrow><a href=# onclick="helpscript(g58);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>����̳�ļ�����Ȩ��
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(48)" type=radio value="1" <%if reGroupSetting(48)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(48)" type=radio value="0" <%if reGroupSetting(48)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g59" value="<b>����̳�ļ�����Ȩ��</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ�����̳�ļ�����Ȩ�ޣ�������Լ�����Ҫ�������ô�ѡ�����԰������������û������ô�Ȩ�ޣ���ع����������̳չ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g59);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<th height="23" colspan="3" align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>���Է��Ͷ���
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(32)" type=radio value="1"  <%if reGroupSetting(32)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(32)" type=radio value="0" <%if reGroupSetting(32)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g60" value="<b>���Է��Ͷ���</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ��п��Է��Ͷ��ŵ�Ȩ�ޣ�������Լ�����Ҫ�������ô�ѡ��">
<td class=Forumrow><a href=# onclick="helpscript(g60);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>��෢���û�
</td>
<td height="23" width="40%" class=ForumrowHighLight><input name="GroupSetting(33)" size=5 type=text value="<%=reGroupSetting(33)%>"></td>
<input type="hidden" id="g61" value="<b>��෢���û�</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û���෢���û���������Լ�����Ҫ�������ô�ѡ����鲻Ҫ���ù�������������̳��Դ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g61);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�������ݴ�С����
</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(34)" size=5 type=text value="<%=reGroupSetting(34)%>"> byte</td>
<input type="hidden" id="g62" value="<b>�������ݴ�С����</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��������ݴ�С���ƣ�������Լ�����Ҫ�������ô�ѡ����鲻Ҫ���ù�������������̳��Դ.">
<td class=Forumrow><a href=# onclick="helpscript(g62);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>�����С����
</td>
<td height="23" width="40%" class=ForumrowHighLight><input name="GroupSetting(35)" size=5 type=text value="<%=reGroupSetting(35)%>"> KB</td>
<input type="hidden" id="g63" value="<b>�����С����</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û������С���ƣ�������Լ�����Ҫ�������ô�ѡ����鲻Ҫ���ù�������������̳��Դ">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g63);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>��ע���û����ٷ��Ӻ���ܷ�����</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(53)" type=text value="<%=reGroupSetting(53)%>" size=4> ����</td>
<input type="hidden" id="g64" value="<b>��ע���û����ٷ��Ӻ���ܷ�����</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ���ע���û����ٷ��Ӻ���ܷ����ţ����ڷ�ֹ����Ⱥ����ʹ�����Ⱥ�����ŵ�Ŀ�ģ�����������ô�ѡ��">
<td class=Forumrow><a href=# onclick="helpscript(g64);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>һ����෢������Ŀ</td>
<td height="23" width="40%" class=ForumrowHighLight><input name="GroupSetting(63)" type=text value="<%=reGroupSetting(63)%>" size=4></td>
<input type="hidden" id="g65" value="<b>һ����෢������Ŀ</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û�һ����෢������Ŀ�����ڷ�ֹ����Ⱥ����ʹ�����Ⱥ�����ŵ�Ŀ�ģ�����������ô�ѡ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g65);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<th height="23" colspan="3" align=left>��������Ȩ��</th>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>����������̳
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(14)" type=radio value="1" <%if reGroupSetting(14)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(14)" type=radio value="0" <%if reGroupSetting(14)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g31" value="<b>����������̳</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ����������̳��������Լ�����Ҫ�������ô�ѡ������δ��¼�û��رմ�ѡ��">
<td class=Forumrow><a href=# onclick="helpscript(g31);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>����ʹ��'���ͱ�ҳ������'����
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(15)" type=radio value="1" <%if reGroupSetting(15)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(15)" type=radio value="0" <%if reGroupSetting(15)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g32" value="<b>����ʹ��'���ͱ�ҳ������'����</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ����ʹ��'���ͱ�ҳ������'���ܣ�������Լ�����Ҫ�������ô�ѡ������δ��¼�û��رմ�ѡ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g32);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�����޸ĸ�������
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(16)" type=radio value="1" <%if reGroupSetting(16)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(16)" type=radio value="0" <%if reGroupSetting(16)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g33" value="<b>�����޸ĸ�������</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ�����޸ĸ�������">
<td class=Forumrow><a href=# onclick="helpscript(g33);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>���������̳�¼�
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(39)" type=radio value="1"  <%if reGroupSetting(39)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(39)" type=radio value="0" <%if reGroupSetting(39)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g34" value="<b>���������̳�¼�</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ���������̳�¼���������Լ�����Ҫ�������ô�ѡ������δ��¼�û��رմ�ѡ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g34);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>�������̳չ����Ȩ��
</td>
<td height="23" width="40%" class=Forumrow>��<input name="GroupSetting(49)" type=radio value="1"  <%if reGroupSetting(49)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(49)" type=radio value="0" <%if reGroupSetting(49)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g35" value="<b>�������̳չ����Ȩ��</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ�������̳չ����Ȩ�ޣ�������Լ�����Ҫ�������ô�ѡ������δ��¼�û��رմ�ѡ��">
<td class=Forumrow><a href=# onclick="helpscript(g35);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=ForumrowHighLight>�Ƿ����ʹ��ǩ��
</td>
<td height="23" width="40%" class=ForumrowHighLight>��<input name="GroupSetting(55)" type=radio value="1"  <%if reGroupSetting(55)="1" then%>checked<%end if%>>&nbsp;��<input name="GroupSetting(55)" type=radio value="0" <%if reGroupSetting(55)="0" then%>checked<%end if%>></td>
<input type="hidden" id="g36" value="<b>�Ƿ����ʹ��ǩ��</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û��Ƿ��Ƿ����ʹ��ǩ����������Լ�����Ҫ�������ô�ѡ��">
<td class=ForumrowHighLight><a href=# onclick="helpscript(g36);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr>
<td height="23" width="60%" class=Forumrow>ǩ������󳤶�</td>
<td height="23" width="40%" class=Forumrow><input name="GroupSetting(56)" type=text value="<%=reGroupSetting(56)%>" size=4> �ֽ�</td>
<input type="hidden" id="g37" value="<b>ǩ������󳤶�</b><br><li>�����������Ը�����Ҫ���ò�ͬ�û����ȼ��û�ǩ������󳤶ȣ�������Լ�����Ҫ�������ô�ѡ�Ϊ�˱���Ӱ���������������������Դ������ù����ֽ���">
<td class=Forumrow><a href=# onclick="helpscript(g37);return false;" class="helplink"><img src="images/manage/help.gif" border=0 title="������Ĺ��������"></a></td>
</tr>
<tr> 
<td height="23" width="60%" class=Forumrow>
</td>
<td height="23" width="40%" class=Forumrow colspan=2><input type="submit" name="submit" value="�� ��"></td>
</tr>
<%
End Function

Function GetGroupPermission()
	Dim i,TempSetting
	For i = 0 To 90
		If Trim(Request.Form("GroupSetting("&i&")"))="" Then
			TempSetting = 0
		Else
			TempSetting = Replace(Trim(Request.Form("GroupSetting("&i&")")),",","")
		End If
		If i = 0 Then
			GetGroupPermission = TempSetting
		ElseIf i = 58 Then
			GetGroupPermission = GetGroupPermission & "," & Replace(Trim(Request.Form("GroupSetting("&i&")A")),",","") & "��" & Replace(Trim(Request.Form("GroupSetting("&i&")B")),",","")
		Else
			GetGroupPermission = GetGroupPermission & "," & TempSetting
		End If
	Next
	GetGroupPermission = Replace(GetGroupPermission,"'","''")
End Function
%>