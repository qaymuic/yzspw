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
'函数名：ReplaceContent
'作  用：文本框中的文本进行排版
'参  数：Content  ----文本字符串
'返回值：处理过的Content
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
'函数名：ReplaceContent
'作  用：文本框中的文本进行排版
'参  数：Content  ----文本字符串
'返回值：处理过的Content
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
'函数名：gotTopic
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str   ----原字符串
'       strlen ----截取长度
'返回值：截取后的字符串
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
			gotTopic=left(str,i) & "…"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

'***********************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
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
'过程名：showpage
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
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
		strTemp=strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>上一页</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "下一页 尾页"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">第" & i & "页</option>"   
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
		strTemp=strTemp & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
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
   	'strTemp=strTemp & "&nbsp;<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
       'strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">第" & i & "页</option>"   
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

'--功能：将阿拉伯数字转换成汉字数字，可选择大小写。
'--参数： 
'engnum	类型：	numeric	传入的数字
'switch	类型：	int	大小写开关：1是小写，2是大写
'返回值		-1：参数engnum非数字行
'		-2：参数switch非法
'---------------------------------------------------------------
function engnum_to_chnnum(engnum,switch)
	'检查参数合法性
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
				case "0" chnnum=chnnum & "零"
				case "1" chnnum=chnnum & "一"
				case "2" chnnum=chnnum & "二"
				case "3" chnnum=chnnum & "三"
				case "4" chnnum=chnnum & "四"
				case "5" chnnum=chnnum & "五"
				case "6" chnnum=chnnum & "六"
				case "7" chnnum=chnnum & "七"
				case "8" chnnum=chnnum & "八"
				case "9" chnnum=chnnum & "九"
				case "." chnnum=chnnum & "点"
				case else 
			end select
		else
			select case ic
				case "0" chnnum=chnnum & "零"
				case "1" chnnum=chnnum & "壹"
				case "2" chnnum=chnnum & "贰"
				case "3" chnnum=chnnum & "叁"
				case "4" chnnum=chnnum & "肆"
				case "5" chnnum=chnnum & "伍"
				case "6" chnnum=chnnum & "陆"
				case "7" chnnum=chnnum & "柒"
				case "8" chnnum=chnnum & "捌"
				case "9" chnnum=chnnum & "玖"
				case "." chnnum=chnnum & "点"
				case else 
			end select
		end if
	next
	engnum_to_chnnum=chnnum
end function
'********************************************
'函数名：IsValidEmail
'作  用：检查Email地址合法性
'参  数：email ----要检查的Email地址
'返回值：True  ----Email地址合法
'       False ----Email地址不合法
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
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'       False ----没有安装
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
'函数名：strLength
'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
'参  数：str  ----要求长度的字符串
'返回值：字符串长度
'**************************************************
function strLength(str)
	ON ERROR RESUME NEXT
	dim WINNT_CHINESE
	WINNT_CHINESE    = (len("中国")=2)
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

'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
	dim fso
	folderpath=Server.MapPath(".")&"\"&folderpath
	Set fso1 = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(FolderPath) then
	'存在
		CheckDir = True
	Else
	'不存在
		CheckDir = False
	End if
	Set fso = nothing
End Function

'-------------根据指定名称生成目录---------
Function MakeNewsDir(foldername)
	dim fso,f
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateFolder(foldername)
    MakeNewsDir = True
	Set fso = nothing
End Function


'****************************************************
'函数名：SendMail
'作  用：用Jmail组件发送邮件
'参  数：ServerAddress  ----服务器地址
'        AddRecipient  ----收信人地址
'        Subject       ----主题
'        Body          ----信件内容
'        Sender        ----发信人地址
'****************************************************
function SendMail(MailtoAddress,MailtoName,Subject,MailBody,FromName,MailFrom,Priority)
	on error resume next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.Message")
	if err then
		SendMail= "<br><li>没有安装JMail组件</li>"
		err.clear
		exit function
	end if
	JMail.Charset="gb2312"          '邮件编码
	JMail.silent=true
	JMail.ContentType = "text/html"     '邮件正文格式
	JMail.ServerAddress=MailServer     '用来发送邮件的SMTP服务器
   	'如果服务器需要SMTP身份验证则还需指定以下参数
	JMail.MailServerUserName = MailServerUserName    '登录用户名
   	JMail.MailServerPassWord = MailServerPassword        '登录密码
  	JMail.MailDomain = MailDomain       '域名（如果用“name@domain.com”这样的用户名登录时，请指明domain.com
	JMail.AddRecipient MailtoAddress,MailtoName     '收信人
	JMail.Subject=Subject         '主题
	JMail.HMTLBody=MailBody       '邮件正文
	JMail.Body="本邮件使用了HTML。"
	JMail.FromName=FromName         '发信人姓名
	JMail.From = MailFrom         '发信人Email
	JMail.Priority=Priority              '邮件等级，1为加急，3为普通，5为低级
	JMail.Send(MailServer)
	SendMail =JMail.ErrorMessage
	JMail.Close
	Set JMail=nothing
end function

'****************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'****************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='title'><td height='22'><strong>错误信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>产生错误的可能原因：</b>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'****************************************************
'过程名：WriteSuccessMsg
'作  用：显示成功提示信息
'参  数：无
'****************************************************
sub WriteSuccessMsg(SuccessMsg)
	dim strSuccess
	strSuccess=strSuccess & "<html><head><title>成功信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strSuccess=strSuccess & "<link href='style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbcrlf
	strSuccess=strSuccess & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center' class='title'><td height='22'><strong>恭喜你！</strong></td></tr>" & vbcrlf
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
		ReplaceBadChar=replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(strChar,"'","‘"),"*","＊"),"?","？"),"(","（"),")","）"),"<","〈"),".","."),">","〉"),",","，"),"[","［"),"]","］"),"!","！"),"&","—"),"%","％"),"#","＃")
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
'函数名：ReplaceContent
'作  用：文本框中的文本进行排版
'参  数：Content  ----文本字符串
'返回值：处理过的Content
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