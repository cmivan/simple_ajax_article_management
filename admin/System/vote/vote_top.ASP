<!--配置文件-->
<!--#include file="../../Bin/Charset.asp"-->
<%Rpath="../../"%>
<!--#include file="../../Bin/adminConn.asp"-->
<!--#include file="../../Bin/Isadmin.asp"-->
<!--#include file="../../Editbox/fckeditor/fckeditor.asp" -->
<!--#include file="../../Class/articles_head.asp"-->
<!--上面的articles_head.asp文件中的样式已经足够，不过为了方便在dreamwear中操纵-->
<link href="images/style.css" rel="stylesheet" type="text/css" />
<link href="../images/style.css" rel="stylesheet" type="text/css" />
<link href="../../images/style.css" rel="stylesheet" type="text/css" />

<script language="javascript">
function test()
{
  if(!confirm('您真的要删除吗？')) return false;
}
</script>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" class="forMy forMAargin">
  <tr>
    <td height="18" class="forumRow">
	|&nbsp;<a href="VOTE_ADMIN.ASP">管理首页</a>
	|&nbsp;<a href="VOTE_ADD.ASP">添加主题</a>
	|</td>
  </tr>
</table>
