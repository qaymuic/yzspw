<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
Dvbbs.Loadtemplates("")
if request("chanWord")="" then Response.End
If Not(Dvbbs.Forum_ChanSetting(0)=1 And Dvbbs.Forum_ChanSetting(7)=1) Then Response.End
if Session("challengeWord")<>Trim(request("chanWord")) then Response.End

Session("challengeWord")=""
Dim rs
set rs=dvbbs.execute("select top 1 * from Dv_ChallengeInfo")
Dim MyForumID
MyForumID=rs("D_ForumID")
set rs=nothing
%>
<form name="redir" action="http://bbs.ray5198.com/addForum.jsp" method="post">
<INPUT type=hidden name="forumId" value="<%=MyForumID%>">
</form>
<script LANGUAGE=javascript>
<!--
redir.submit();
//-->
</script>
