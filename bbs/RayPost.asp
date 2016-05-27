<%@ LANGUAGE = VBScript CodePage = 936%>
<%
Option Explicit
Dim FrameBody
%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
If Request.QueryString("PostType")=0 Then
'http://bbs.ray5198.com/SearchEngine/searchforum.jsp
	FrameBody="<form name=""redir"" action=""http://www.findbbs.com/SearchEngine/searchforum.jsp"" method=""post"">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""forumname"" value="""&Server.htmlencode(Request.QueryString("forumname"))&""">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""forumurl"" value="""&Request.QueryString("forumurl")&""">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""forumlogincount"" value="""&Request.QueryString("forumlogincount")&""">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""foruminlinecount"" value="""&Request.QueryString("foruminlinecount")&""">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""forumtitlecount"" value="""&Request.QueryString("forumtitlecount")&""">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""forumvisitprob"" value=""1"">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""forumemail"" value="""&Request.QueryString("forumemail")&""">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""forumtag"" value=""host"">"
	FrameBody=FrameBody+"</form>"
ElseIf Request.QueryString("PostType")=1 Then
'http://bbs.ray5198.com/SearchEngine/searchforumchild.jsp
	FrameBody=FrameBody+"<form name=""redir"" action=""http://www.findbbs.com/SearchEngine/searchforumchild.jsp"" method=""post"">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""forumchildname"" value="""&Server.htmlencode(Request.QueryString("forumchildname"))&""">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""forumchildurl"" value="""&Request.QueryString("forumchildurl")&""">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""forumchildtitlecount"" value="""&Request.QueryString("forumchildtitlecount")&""">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""foruminlinecount"" value="""&Request.QueryString("foruminlinecount")&""">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""forumlogincount"" value="""&Request.QueryString("forumlogincount")&""">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""forumvisitprob"" value=""1"">"
	FrameBody=FrameBody+"<INPUT type=hidden name=""forumchildtag"" value=""subjection"">"
	FrameBody=FrameBody+"</form>"
End If

If IsNumeric(Request.QueryString("PostType")) Then
Response.Write FrameBody
%>
<script LANGUAGE=javascript>
redir.submit();
</script>
<%End If%>