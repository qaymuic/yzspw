<%
'=================================================
'����Dv7.0Jmail�����ʼ������
'   Edit By YangZheng
'=================================================
Dim SendMail

Sub Jmail(Email,Topic,Mailbody)
	On Error Resume Next
	Dim JMail
	Set JMail = Server.CreateObject("JMail.Message")
	JMail.silent=true
	JMail.Logging = True
	JMail.Charset = "gb2312"
	If Not(Dvbbs.Forum_info(12) = "" Or Dvbbs.Forum_info(13) = "") Then
		JMail.MailServerUserName = Dvbbs.Forum_info(12) '�����ʼ���������¼��
		JMail.MailServerPassword = Dvbbs.Forum_info(13) '��¼����
	End If
	JMail.ContentType = "text/html"
	JMail.Priority = 1
	JMail.From = Dvbbs.Forum_info(5)
	JMail.FromName = Dvbbs.Forum_info(0)
	JMail.AddRecipient Email
	JMail.Subject = Topic
	JMail.Body = Mailbody
	JMail.Send (Dvbbs.Forum_info(4))
	Set JMail = Nothing
	SendMail = "OK"
	If Err Then SendMail = "False"
End Sub
	
Sub Cdonts(Email,Topic,Mailbody)
	On Error Resume Next
	Dim ObjCDOMail
	Set ObjCDOMail = Server.CreateObject("CDONTS.NewMail")
	ObjCDOMail.From = Dvbbs.Forum_info(5)
	ObjCDOMail.To = Email
	ObjCDOMail.Subject = Topic
	ObjCDOMail.BodyFormat = 0 
	ObjCDOMail.MailFormat = 0 
	ObjCDOMail.Body = Mailbody
	ObjCDOMail.Send
	Set ObjCDOMail = Nothing
	SendMail = "OK"
	If Err Then SendMail = "False"
End Sub

Sub Aspemail(Email,Topic,Mailbody)
	On Error Resume Next
	Dim Mailer
	Set Mailer = Server.CreateObject("Persits.MailSender")
	Mailer.Charset = "gb2312"
	Mailer.IsHTML = True
	Mailer.username = Dvbbs.Forum_info(12)	'����������Ч���û���
	Mailer.password = Dvbbs.Forum_info(13)	'����������Ч������
	Mailer.Priority = 1
	Mailer.Host = Dvbbs.Forum_info(4)
	Mailer.Port = 25 ' �����ѡ.�˿�25��Ĭ��ֵ
	Mailer.From = Dvbbs.Forum_info(5)
	Mailer.FromName = Dvbbs.Forum_info(0) ' �����ѡ
	Mailer.AddAddress Email,Email
	Mailer.Subject = Topic
	Mailer.Body = Mailbody
	Mailer.Send
	SendMail = "OK"
	If Err Then SendMail = "False"
End Sub
%>