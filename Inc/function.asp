<%
						function nohtml(str)
							dim re
							Set re=new RegExp
							re.IgnoreCase =true
							re.Global=True
							re.Pattern="(\<.[^\<]*\>)"
							str=re.replace(str," ")
							re.Pattern="(\<\/[^\<]*\>)"
							str=re.replace(str," ")
							nohtml=str
							set re=nothing
						end function
'***********************************************
'��������ReplaceContent
'��  �ã��ı����е��ı������Ű�
'��  ����Content  ----�ı��ַ���
'����ֵ���������Content
'***********************************************
Function ReplaceContent(Content)
	Content = replace(Content, ">", "&gt;")
    Content = replace(Content, "<", "&lt;")

    Content = Replace(Content, CHR(32), "&nbsp;")
    Content = Replace(Content, CHR(9), "&nbsp;")
    Content = Replace(Content, CHR(34), "&quot;")
    Content = Replace(Content, CHR(39), "&#39;")
    Content = Replace(Content, CHR(13), "")
    Content = Replace(Content, CHR(10) & CHR(10), "</P><P> ")
    Content = Replace(Content, CHR(10), "<BR> ")
	ReplaceContent = Content
End Function
'***********************************************
'��������ReplaceContent
'��  �ã��ı����е��ı������Ű�
'��  ����Content  ----�ı��ַ���
'����ֵ���������Content
'***********************************************
Function ReReplaceContent(Content)
	Content = replace(Content, "&gt;", ">")
    Content = replace(Content, "&lt;", "<")

    Content = Replace(Content, "&nbsp;", CHR(32))
    Content = Replace(Content, "&nbsp;", CHR(9))
    Content = Replace(Content, "&quot;", CHR(34))
    Content = Replace(Content, "&#39;", CHR(39))
    Content = Replace(Content, "",CHR(13) )
    Content = Replace(Content, "</P><P> ", CHR(10) & CHR(10))
    Content = Replace(Content, "<BR> ", CHR(10))
	ReReplaceContent = Content
End Function

function getFileExtName(fileName)
dim pos
pos=instrrev(filename,".")
if pos>0 then
getFileExtName=mid(fileName,pos+1)
else
getFileExtName=""
end if
end function
'*************************************************
'��������gotTopic
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str   ----ԭ�ַ���
'       strlen ----��ȡ����
'����ֵ����ȡ����ַ���
'*************************************************
function gotTopic(str,strlen)
	if str="" then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i) & "��"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

'***********************************************
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
'***********************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function

'***********************************************
'��������showpage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       totalnumber ----������
'       maxperpage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
'***********************************************
sub showpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= " <br><table width='98%'  bgcolor='#F5F5F5'><tr><td><div align='right'>"
	if ShowTotal=true then 
		strTemp=strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "��ҳ ��һҳ&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>��һҳ</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "��һҳ βҳ"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>��һҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>βҳ</a>"
  	end if
   	strTemp=strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;ת����<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">��" & i & "ҳ</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</div></td></tr></table>"
	response.write strTemp
end sub

sub showpage1(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table width='98%' align='center'><tr><td><div align='right'>"
	'if ShowTotal=true then 
		strTemp=strTemp & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	'end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & " <font face=webdings  color=black>7</font>&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'></a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'><font face=webdings  color=black>7</font></a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "<font face=webdings  color=black>8</font> "
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'><font face=webdings  color=black>8</font></a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'></a>"
  	end if
   	'strTemp=strTemp & "&nbsp;<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
       'strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;ת����<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">��" & i & "ҳ</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</div></td></tr></table>"
	response.write strTemp
end sub

function crq(rq)
rq=year(rq)&"-"&month(rq)&"-"&day(rq)
crq=rq
end function

'--���ܣ�������������ת���ɺ������֣���ѡ���Сд��
'--������ 
'engnum	���ͣ�	numeric	���������
'switch	���ͣ�	int	��Сд���أ�1��Сд��2�Ǵ�д
'����ֵ		-1������engnum��������
'		-2������switch�Ƿ�
'---------------------------------------------------------------
function engnum_to_chnnum(engnum,switch)
	'�������Ϸ���
	if not isnumeric(engnum) then
		engnum_to_chnnum="-1"
		exit function
	end if
	if not (cint(switch)=1 or cint(switch)=2) then
		engnum_to_chnnum="-2"
		exit function
	end if
	dim strengnum,lenengnum,chnnum
	dim i,ic
	strengnum=cstr(engnum)
	lenengnum=len(strengnum)
	for i=1 to lenengnum
		ic=mid(strengnum,i,1)
		if cint(switch)=1 then
			select case ic
				case "0" chnnum=chnnum & "��"
				case "1" chnnum=chnnum & "һ"
				case "2" chnnum=chnnum & "��"
				case "3" chnnum=chnnum & "��"
				case "4" chnnum=chnnum & "��"
				case "5" chnnum=chnnum & "��"
				case "6" chnnum=chnnum & "��"
				case "7" chnnum=chnnum & "��"
				case "8" chnnum=chnnum & "��"
				case "9" chnnum=chnnum & "��"
				case "." chnnum=chnnum & "��"
				case else 
			end select
		else
			select case ic
				case "0" chnnum=chnnum & "��"
				case "1" chnnum=chnnum & "Ҽ"
				case "2" chnnum=chnnum & "��"
				case "3" chnnum=chnnum & "��"
				case "4" chnnum=chnnum & "��"
				case "5" chnnum=chnnum & "��"
				case "6" chnnum=chnnum & "½"
				case "7" chnnum=chnnum & "��"
				case "8" chnnum=chnnum & "��"
				case "9" chnnum=chnnum & "��"
				case "." chnnum=chnnum & "��"
				case else 
			end select
		end if
	next
	engnum_to_chnnum=chnnum
end function
'********************************************
'��������IsValidEmail
'��  �ã����Email��ַ�Ϸ���
'��  ����email ----Ҫ����Email��ַ
'����ֵ��True  ----Email��ַ�Ϸ�
'       False ----Email��ַ���Ϸ�
'********************************************
function IsValidEmail(email)
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
	   IsValidEmail = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmail = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmail = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmail = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmail = false
	   exit function
	end if
	if InStr(email, "..") > 0 then
	   IsValidEmail = false
	end if
end function

'***************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'       False ----û�а�װ
'***************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

'**************************************************
'��������strLength
'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��  ����str  ----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
'**************************************************
function strLength(str)
	ON ERROR RESUME NEXT
	dim WINNT_CHINESE
	WINNT_CHINESE    = (len("�й�")=2)
	if WINNT_CHINESE then
        dim l,t,c
        dim i
        l=len(str)
        t=l
        for i=1 to l
        	c=asc(mid(str,i,1))
            if c<0 then c=c+65536
            if c>255 then
                t=t+1
            end if
        next
        strLength=t
    else 
        strLength=len(str)
    end if
    if err.number<>0 then err.clear
end function

'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
	dim fso
	folderpath=Server.MapPath(".")&"\"&folderpath
	Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(FolderPath) then
	'����
		CheckDir = True
	Else
	'������
		CheckDir = False
	End if
	Set fso = nothing
End Function

'-------------����ָ����������Ŀ¼---------
Function MakeNewsDir(foldername)
	dim fso,f
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateFolder(foldername)
    MakeNewsDir = True
	Set fso = nothing
End Function


'****************************************************
'��������SendMail
'��  �ã���Jmail��������ʼ�
'��  ����ServerAddress  ----��������ַ
'        AddRecipient  ----�����˵�ַ
'        Subject       ----����
'        Body          ----�ż�����
'        Sender        ----�����˵�ַ
'****************************************************
function SendMail(MailtoAddress,MailtoName,Subject,MailBody,FromName,MailFrom,Priority)
	on error resume next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.Message")
	if err then
		SendMail= "<br><li>û�а�װJMail���</li>"
		err.clear
		exit function
	end if
	JMail.Charset="gb2312"          '�ʼ�����
	JMail.silent=true
	JMail.ContentType = "text/html"     '�ʼ����ĸ�ʽ
	JMail.ServerAddress=MailServer     '���������ʼ���SMTP������
   	'�����������ҪSMTP�����֤����ָ�����²���
	JMail.MailServerUserName = MailServerUserName    '��¼�û���
   	JMail.MailServerPassWord = MailServerPassword        '��¼����
  	JMail.MailDomain = MailDomain       '����������á�name@domain.com���������û�����¼ʱ����ָ��domain.com
	JMail.AddRecipient MailtoAddress,MailtoName     '������
	JMail.Subject=Subject         '����
	JMail.HMTLBody=MailBody       '�ʼ�����
	JMail.Body="���ʼ�ʹ����HTML��"
	JMail.FromName=FromName         '����������
	JMail.From = MailFrom         '������Email
	JMail.Priority=Priority              '�ʼ��ȼ���1Ϊ�Ӽ���3Ϊ��ͨ��5Ϊ�ͼ�
	JMail.Send(MailServer)
	SendMail =JMail.ErrorMessage
	JMail.Close
	Set JMail=nothing
end function

'****************************************************
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'****************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>��������Ŀ���ԭ��</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'****************************************************
'��������WriteSuccessMsg
'��  �ã���ʾ�ɹ���ʾ��Ϣ
'��  ������
'****************************************************
sub WriteSuccessMsg(SuccessMsg)
	dim strSuccess
	strSuccess=strSuccess & "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strSuccess=strSuccess & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strSuccess=strSuccess & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center' class='title'><td height='22'><strong>��ϲ�㣡</strong></td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr class='tdbg'><td height='100' valign='top'><br>" & SuccessMsg &"</td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center' class='tdbg'><td>&nbsp;</td></tr>" & vbcrlf
	strSuccess=strSuccess & "</table>" & vbcrlf
	strSuccess=strSuccess & "</body></html>" & vbcrlf
	response.write strSuccess
end sub

function ReplaceBadChar(strChar)
	if strChar="" then
		ReplaceBadChar=""
	else
		ReplaceBadChar=replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(strChar,"'","��"),"*","��"),"?","��"),"(","��"),")","��"),"<","��"),".","."),">","��"),",","��"),"[","��"),"]","��"),"!","��"),"&","��"),"%","��"),"#","��")
	end if
end function

function dvHTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")

    dvHTMLEncode = fString
end if
end function

'***********************************************
'��������ReplaceContent
'��  �ã��ı����е��ı������Ű�
'��  ����Content  ----�ı��ַ���
'����ֵ���������Content
'***********************************************
Function ReplaceContent(Content)
	Content = replace(Content, ">", "&gt;")
    Content = replace(Content, "<", "&lt;")

    Content = Replace(Content, CHR(32), "&nbsp;")
    Content = Replace(Content, CHR(9), "&nbsp;")
    Content = Replace(Content, CHR(34), "&quot;")
    Content = Replace(Content, CHR(39), "&#39;")
    Content = Replace(Content, CHR(13), "")
    Content = Replace(Content, CHR(10) & CHR(10), "</P><P> ")
    Content = Replace(Content, CHR(10), "<BR> ")
	ReplaceContent = Content
End Function
%>