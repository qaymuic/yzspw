<!--#include file =conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/DvADChar.asp" -->
<%	
	Head()
	dim admin_flag,rs_c
	admin_flag=",1,"
	If Not Dvbbs.master or instr(","&session("flag")&",",admin_flag)=0 Then 
		Errmsg=ErrMsg + "<BR><li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		Call dvbbs_error()
	Else
		if request("action")="save" then
		call saveconst()
		elseif request("action")="restore" then
		call restore()
		else
		call consted()
		end if
		Footer()
	end if


Sub consted()
Dim  sel
%>
<form method="POST" action="admin_Challenge.asp?action=save">
<table width="95%" border="0" cellspacing="0" cellpadding="3"  align=center class="tableBorder"> 
<th height=25 colspan=2 align=center id=tabletitlelink><a name="setting20"></a><b>��̳���Ż�����Ϣ����</b>[<a href="admin_Challenge.asp?action=restore">��ԭĬ������</a>]
</tr>
<tr> 
<td width="50%" class=Forumrow> <U>�Ƿ����������</U><br>��������ܿ��أ�ѡ�������������������þ���Ч��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(0)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(0))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_ChanSetting(0)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(0))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<td width="50%" class=Forumrow> <U>�Ƿ���������</U><br>�������ܿ��أ�ѡ��������⻥����������ʾ��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(1)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(1))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_ChanSetting(1)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(1))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>����ͨ��ģʽ</U><br>���⻥�����������ʾģʽ��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(2)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(2))=0 then%>checked<%end if%>>����ʾ&nbsp;
<input type=radio name="Forum_ChanSetting(2)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(2))=1 then%>checked<%end if%>>��ʾ&nbsp;
<input type=radio name="Forum_ChanSetting(2)" value=2 <%if cint(Dvbbs.Forum_ChanSetting(2))=2 then%>checked<%end if%>>��ʾ�������˵���&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>����bannerģʽ</U><br>��̳��������ʾ�Ĺ�档</td>
<td width="50%" class=Forumrow>
<input type=radio name="Forum_ChanSetting(3)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(3))=0 then%>checked<%end if%>>��̳���&nbsp;
<input type=radio name="Forum_ChanSetting(3)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(3))=1 then%>checked<%end if%>>���Ź��&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>β��ͨ��ģʽ</U><br>��̳�ײ���ʾ�Ĺ�档</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(4)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(4))=0 then%>checked<%end if%>>��̳���&nbsp;
<input type=radio name="Forum_ChanSetting(4)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(4))=1 then%>checked<%end if%>>���Ź��&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>������ģʽ</U><br>���Ӽ��Ƿ���ʾ���Ź�档</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(5)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(5))=0 then%>checked<%end if%>>��&nbsp;(��ʾ��̳������)
<input type=radio name="Forum_ChanSetting(5)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(5))=1 then%>checked<%end if%>>��&nbsp;(��ʾ����������)
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>�Ƿ���վ�ڶ��Ż���</U><br>��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(6)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(6))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_ChanSetting(6)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(6))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>�Ƿ������ⶩ��</U><br>��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(7)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(7))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_ChanSetting(7)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(7))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>�Ƿ���VIP��̳</U><br>��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(8)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(8))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_ChanSetting(8)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(8))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>�Ƿ����������Աע�ᡢ�޸�����</U><br>��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(9)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(9))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_ChanSetting(9)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(9))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>�Ƿ������ޱ߽��¼</U><br>��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(10)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(10))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_ChanSetting(10)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(10))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>�Ƿ���ͬ����湦��</U><br>��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(11)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(11))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_ChanSetting(11)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(11))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
</tr>
<td width="50%" class=Forumrow> <U>��¼��ע��ɹ��Ƿ���������Ա���</U><br>��</td>
<td width="50%" class=Forumrow>  
<input type=radio name="Forum_ChanSetting(12)" value=0 <%if cint(Dvbbs.Forum_ChanSetting(12))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_ChanSetting(12)" value=1 <%if cint(Dvbbs.Forum_ChanSetting(12))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>

<tr>
<td width="50%" class=Forumrow> &nbsp;</td>
<td width="50%" class=Forumrow>  
<div align="center"> 
<input type="submit" name="Submit" value="�� ��">
</div>
</td>
</tr>
</table>
</form>
<%
end sub

sub saveconst()
dim Forum_ChanSetting,ChanSetting

'Forum_ChanSetting=request.form("Forum_ChanSetting(0)") & "," & request.form("Forum_ChanSetting(1)") & "," & request.form("Forum_ChanSetting(2)") & "," & request.form("Forum_ChanSetting(3)") & "," & request.form("Forum_ChanSetting(4)") & "," & request.form("Forum_ChanSetting(5)") & "," & request.form("Forum_ChanSetting(6)") & "," & request.form("Forum_ChanSetting(7)") & "," & request.form("Forum_ChanSetting(8)") & "," & request.form("Forum_ChanSetting(9)") & "," & request.form("Forum_ChanSetting(10)") & "," & request.form("Forum_ChanSetting(11)") & "," & request.form("Forum_ChanSetting(12)")
For i=0 To 60
	If Request.Form("Forum_ChanSetting("&i&")")="" Then
		ChanSetting = 1
	Else
		ChanSetting = Replace(Request.Form("Forum_ChanSetting("&i&")"),",","")
	End If
	If i = 0 Then
		Forum_ChanSetting = ChanSetting
	Else
		Forum_ChanSetting = Forum_ChanSetting & "," & ChanSetting
	End If
Next

sql="update Dv_setup set Forum_ChanSetting='"&Forum_ChanSetting&"'"
dvbbs.execute(sql)
Dvbbs.Name="setup"
Dvbbs.reloadsetup()
Dv_suc("���ö��Ż������ܳɹ�")

end sub

'�ָ�Ĭ������
Sub restore()
	Dvbbs.Forum_ChanSetting="1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1"
	Dvbbs.Execute("update Dv_setup set Forum_ChanSetting='"&dvbbs.Forum_ChanSetting&"'")
	Dv_suc("��ԭ����")
End Sub 
%>